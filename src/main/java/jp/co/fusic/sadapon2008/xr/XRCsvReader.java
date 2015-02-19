package jp.co.fusic.sadapon2008.xr;

import java.io.FileInputStream;
import java.io.InputStreamReader;
import java.io.IOException;
import java.util.ArrayList;
import au.com.bytecode.opencsv.CSVReader;

public class XRCsvReader {
    public ArrayList<ArrayList<String>> readFile(String filename) throws InvalidCsvFileXRException, IOException {
        ArrayList<ArrayList<String>> result = new ArrayList<ArrayList<String>>();
        try {
            FileInputStream input= new FileInputStream(filename);
            InputStreamReader inReader=new InputStreamReader(input, "UTF-8");
            CSVReader reader = new CSVReader(inReader,',','"');
            String [] nextLine;

            // 1～2行目を読み込む
            for(int i = 1; i <= 2; i++ ) {
                if((nextLine = reader.readNext()) == null) {
                    throw new InvalidCsvFileXRException("error at line " + Integer.toString(i));
                }

                ArrayList<String> line_frist = new ArrayList<String>();
                for(String val: nextLine) {
                    line_frist.add(val);
                }

                result.add(line_frist);
            }
            reader.close();
        } catch(IOException e) {
            throw e;
        }

        return result;
    }
}
