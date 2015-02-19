package jp.co.fusic.sadapon2008.xr;

import java.util.ArrayList;
import java.util.Map;
import java.util.HashMap;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.awt.Dimension;
import java.io.IOException;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFCreationHelper;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.util.IOUtils;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class XRXlsRenderer {

    private static final String valuePrefixPattern = "^(#|＃)";

    // 1行目をキー、2行目を値とするマップに変換する
    private Map<String,String> combineKeyValue(ArrayList<ArrayList<String>> data) {
        Map<String,String> result = new HashMap<String,String>();

        int size = data.get(1).size();

        int i = 0;
        for(String key: data.get(0)) {
            String val = null;
            if(size-1 >= i) {
                val = data.get(1).get(i);
            }
            result.put(key, val);
            i++;
        }

        return result;
    }

    // 画像を指定されたセルの矩形領域に縦横比をなるべく維持して引き伸ばして中央ぞろえにして挿入する
    protected void setPictureSize(HSSFPicture picture, int pixcels_per_col, int pixcels_per_row, int col, int row, int num_col, int num_row) {
        Dimension d = picture.getImageDimension();
        HSSFClientAnchor anchor = picture.getClientAnchor();

        int target_width = num_col * pixcels_per_col;
        int target_height = num_row * pixcels_per_row;

        if((d.width == 0) || (target_width == 0)) {
            return;
        }

        if((double)d.height/(double)d.width > (double)target_height/(double)target_width) {
            // 高さでそろえる
            anchor.setRow2(row + (num_row-1));
            anchor.setDy2(255);

            int new_width = (int)Math.round((double)(d.width * target_height)/(double)d.height);
            if(new_width > target_width) {
                new_width = target_width;
            }

            int diff_width = target_width - new_width;
            int offset_x1 = diff_width/2;
            int offset_x2 = offset_x1 + diff_width%2;

            int col1 = col + offset_x1/pixcels_per_col;
            anchor.setCol1(col1);
            anchor.setDx1(offset_x1%pixcels_per_col * (1024/pixcels_per_col));

            int col2 = col + (num_col-1) - offset_x2/pixcels_per_col;
            anchor.setCol2(col2);
            anchor.setDx2(1023 - (offset_x2%pixcels_per_col * (1024/pixcels_per_col)));
        } else {
            // 幅でそろえる
            anchor.setCol2(col + (num_col-1));
            anchor.setDx2(1023);

            int new_height = (int)Math.round((double)(d.height * target_width)/(double)d.width);
            if(new_height > target_height) {
                new_height = target_height;
            }
            int diff_height = target_height - new_height;
            int offset_y1 = diff_height/2;
            int offset_y2 = offset_y1 + diff_height%2;

            int row1 = row + offset_y1/pixcels_per_row;
            anchor.setRow1(row1);
            anchor.setDy1(offset_y1%pixcels_per_row * (256/pixcels_per_row));

            int row2 = row + (num_row-1) - offset_y2/pixcels_per_row;
            anchor.setRow2(row2);
            anchor.setDy2(255 - (offset_y2%pixcels_per_row * (256/pixcels_per_row)));
        }
    }

    public void renderFile(ArrayList<ArrayList<String>> data, ArrayList<ArrayList<String>> parameter, String filename_template, String filename_output, int pixcels_per_col, int pixcels_per_row) throws IOException {
        // データとパラメータのmapを準備する
        Map<String,String> map_data = combineKeyValue(data);
        Map<String,String> map_parameter = combineKeyValue(parameter);

        // テンプレートファイルを開く
        HSSFWorkbook book;

        try {
            FileInputStream is = new FileInputStream(filename_template);
            book = new HSSFWorkbook(is);
        } catch(IOException e) {
            throw e;
        }

        // シートごとに処理する
        for(int i = 0; i < book.getNumberOfSheets(); i++) {
            HSSFSheet sheet = book.getSheetAt(i);

            // 行ごとに処理する
            for(int r = sheet.getFirstRowNum(); r <= sheet.getLastRowNum(); r++) {
                HSSFRow row = sheet.getRow(r);
                if(row == null) {
                    continue;
                }

                // セルごとに処理する
                for(int c = row.getFirstCellNum(); c <= row.getLastCellNum(); c++) {
                    HSSFCell cell = row.getCell(c);

                    // 値が文字列でなければスキップする
                    if((cell == null) || (cell.getCellType() != Cell.CELL_TYPE_STRING)) {
                        continue;
                    }

                    // 値が定義された文字列で始まらない場合はスキップする
                    String cell_val = cell.getStringCellValue();
                    if(cell_val == null) {
                        continue;
                    }

                    Pattern p = Pattern.compile(valuePrefixPattern);
                    Matcher m = p.matcher(cell_val);
                    if(!m.find()) {
                        continue;
                    }

                    // 置換するキーを取得
                    String key = cell_val.replaceAll(valuePrefixPattern, "");

                    // とりあえず空白にする
                    cell.setCellValue("");

                    // データが存在しない場合はスキップする
                    if(!map_data.containsKey(key)) {
                        continue;
                    }

                    String val = map_data.get(key);
                    if(val == null) {
                        continue;
                    }

                    // パラメータをチェックする
                    String[] params = {};
                    if(map_parameter.containsKey(key)) {
                        // パラメータはカンマ区切りで分割する
                        params = map_parameter.get(key).split(",");
                    }

                    if((params == null) || (params.length < 1) || ("string".equalsIgnoreCase(params[0]))) {
                        // 文字列
                        cell.setCellValue(val);
                    } else if("numeric".equalsIgnoreCase(params[0])) {
                        // 数値
                        double n;
                        try {
                            n = Double.parseDouble(val);
                        } catch(NumberFormatException e) {
                            // エラー
                            continue;
                        }
                        cell.setCellType(Cell.CELL_TYPE_NUMERIC);
                        cell.setCellValue(n);
                    } else if("formula".equalsIgnoreCase(params[0])) {
                        // 式
                        cell.setCellType(Cell.CELL_TYPE_FORMULA);
                        cell.setCellValue(val);
                    } else if((params.length >= 3) && (("image_jpg".equalsIgnoreCase(params[0])) || ("image_png".equalsIgnoreCase(params[0])))) {
                        // 画像(jpg,png)
                        int num_col;
                        int num_row;
                        try {
                            num_col = Integer.parseInt(params[1]);
                        } catch(NumberFormatException e) {
                            // エラー
                            continue;
                        }
                        try {
                            num_row = Integer.parseInt(params[2]);
                        } catch(NumberFormatException e) {
                            // エラー
                            continue;
                        }

                        // 画像ファイルを開いて追加する
                        int pictureIndex;
                        try {
                            FileInputStream img = new FileInputStream(val);
                            byte[] bytes = IOUtils.toByteArray(img);
                            if("image_jpg".equalsIgnoreCase(params[0])) {
                                pictureIndex = book.addPicture(bytes, HSSFWorkbook.PICTURE_TYPE_JPEG);
                            } else {
                                pictureIndex = book.addPicture(bytes, HSSFWorkbook.PICTURE_TYPE_PNG);
                            }
                            img.close();
                        } catch(IOException e) {
                            continue;
                        }

                        // 画像を挿入する
                        HSSFCreationHelper helper = (HSSFCreationHelper) book.getCreationHelper();
                        HSSFPatriarch patriarch = sheet.createDrawingPatriarch();
                        HSSFClientAnchor clientAnchor = helper.createClientAnchor();
                        clientAnchor.setAnchorType(ClientAnchor.DONT_MOVE_AND_RESIZE);
                        clientAnchor.setCol1(c);
                        clientAnchor.setRow1(r);
                        clientAnchor.setDx1(0);
                        clientAnchor.setDy1(0);
                        clientAnchor.setDx2(1023);
                        clientAnchor.setDy2(255);
                        HSSFPicture picture = patriarch.createPicture(clientAnchor, pictureIndex);
                        picture.resize();

                        // サイズと位置を調整する
                        setPictureSize(picture, pixcels_per_col, pixcels_per_row, c, r, num_col, num_row);
                    }
                }
            }
        }

        // ファイルを出力する
        try {
            FileOutputStream os = new FileOutputStream(filename_output);
            book.write(os);
        } catch(IOException e) {
            throw e;
        }
    }
}
