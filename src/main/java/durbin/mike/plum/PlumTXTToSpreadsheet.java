package durbin.mike.plum;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.UnsupportedEncodingException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.swing.JOptionPane;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;

public class PlumTXTToSpreadsheet {

    public static final char FIELD_DELIMITER = (char) 253;

    public static final Pattern FIELD_PATTERN = Pattern.compile("([A-Z0-9]{3})(.*)");

    public static void main(String [] args) throws IOException {

        File sourceDir = new File(args[0]);
        File outputDir = new File(args[1]);
        for (File f : sourceDir.listFiles()) {
            if (f.getName().endsWith(".txt")) {
                convertText(f);
            }
        }
    }

    public static void convertText(File textFile) throws IOException {
        FileInputStream fis = new FileInputStream(textFile);
        BufferedReader r = new BufferedReader(new InputStreamReader(fis, "ISO-8859-1"));
    
        Workbook wb = new HSSFWorkbook();
    
        Map<List<String>, Sheet> sheetMap = new HashMap<List<String>, Sheet>();
    
        int count = 0;
        String line = null;
        while ((line = r.readLine()) != null) {
            List<String> c = new ArrayList<String>();
            c.add("original order");
            List<String> values = new ArrayList<String>();
            values.add(String.valueOf(count));
            count ++;
            for (String field : line.split("\\Q" + FIELD_DELIMITER + "\\E")) {
                Matcher m = FIELD_PATTERN.matcher(field);
                if (m.matches()) {
                    c.add(m.group(1));
                    values.add(m.group(2));
                } else {
                    throw new RuntimeException("\"" + field + "\" does not have a field code!");
                }
            }
                
            Sheet s = sheetMap.get(c);
            if (s == null) {
                s = wb.createSheet(WorkbookUtil.createSafeSheetName("Type " + sheetMap.size()));
                appendRow(s, c);
                sheetMap.put(c, s);
            }
            appendRow(s, values);
        }
        String outputFilename = textFile.getName().replace(".txt", "") + ".xls";
        while (new File(outputFilename).exists()) {
            DateFormat f = new SimpleDateFormat("yyyy-MMM-dd.HH.mm.ss");
            outputFilename = outputFilename.replace(".xls", "-" + f.format(new Date()) + ".xls");
        }
        FileOutputStream fos = new FileOutputStream(new File(outputFilename));
        wb.write(fos);
        fos.close();

        JOptionPane.showMessageDialog(null, "Created Spreadsheet file (" + count + " records): " + new File(outputFilename).getAbsolutePath());
    }
    
    private static void appendRow(Sheet s, List<String> values) {
        Row r = s.createRow(s.getLastRowNum() + 1);
        //System.out.println("Created row "  + r.getRowNum());
        for (int i = 0; i < values.size(); i ++) {
            r.createCell(i).setCellValue(values.get(i));
        }
    }
}
