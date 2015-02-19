package jp.co.fusic.sadapon2008.xr;

import java.io.IOException;
import java.util.ArrayList;

import org.apache.commons.cli.BasicParser;
import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.ParseException;

import org.apache.commons.lang.math.NumberUtils;

public class App {
    public static void main(String[] args) {

        String filename_data_csv = "";
        String filename_parameter_csv = "";
        String filename_template_xls = "";
        String filename_output_xls = "";
        String pixcel_per_col = "";
        String pixcel_per_row = "";

        // コマンドライン引数を解析する
        Options opts = new Options();
        opts.addOption("d", true, "input data csv file");
        opts.addOption("p", true, "input parameter csv file");
        opts.addOption("t", true, "input template xls file");
        opts.addOption("o", true, "output xls file");
        opts.addOption("pc", true, "number of pixcels per col for image (default: 8)");
        opts.addOption("pr", true, "number of pixcels per row for image (default: 8)");
        BasicParser parser = new BasicParser();
        CommandLine cl;
        HelpFormatter help = new HelpFormatter();

        try {
            // 解析する
            cl = parser.parse(opts, args);

            filename_data_csv = cl.getOptionValue("d");
            filename_parameter_csv = cl.getOptionValue("p");
            filename_template_xls = cl.getOptionValue("t");
            filename_output_xls = cl.getOptionValue("o");

            pixcel_per_col = cl.getOptionValue("pc", "8");
            pixcel_per_row = cl.getOptionValue("pr", "8");

            if((filename_data_csv == null)
               || (filename_parameter_csv == null)
               || (filename_template_xls == null)
               || (filename_output_xls == null)) {
                throw new ParseException("");
            }

            if(!NumberUtils.isDigits(pixcel_per_col) || !NumberUtils.isDigits(pixcel_per_row)) {
                throw new ParseException("");
            }
        } catch(ParseException e) {
            help.printHelp("xsl_renderer", opts);
            System.exit(1);
        }

        // CSVファイルを読み込む
        XRCsvReader csv_reader = new XRCsvReader();
        ArrayList<ArrayList<String>> data_result = null;
        ArrayList<ArrayList<String>> parameter_result = null;
        try {
            data_result = csv_reader.readFile(filename_data_csv);
        } catch(IOException e) {
            System.err.println("error: failed to open data csv file");
            System.exit(1);
        } catch(InvalidCsvFileXRException e) {
            System.err.println("error: invalid data csv file");
            System.exit(1);
        }
        try {
            parameter_result = csv_reader.readFile(filename_parameter_csv);
        } catch(IOException e) {
            System.err.println("error: failed to open parameter csv file");
            System.exit(1);
        } catch(InvalidCsvFileXRException e) {
            System.err.println("error: invalid parameter csv file");
            System.exit(1);
        }

        try {
            XRXlsRenderer renderer = new XRXlsRenderer();
            renderer.renderFile(data_result, parameter_result, filename_template_xls, filename_output_xls, Integer.parseInt(pixcel_per_col), Integer.parseInt(pixcel_per_row));
        } catch(IOException e) {
            System.err.println("error: failed to access excel file");
            System.exit(1);
        }
    }
}
