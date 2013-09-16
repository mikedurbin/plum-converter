package durbin.mike.plum;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.PrintWriter;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PlumSpreadsheetToTXT {

    public static final char FIELD_DELIMITER = (char) 253;

    public static void main(String [] args) throws IOException {
        File[] files = null;
        if (args.length == 0) {
            JFileChooser chooser = new JFileChooser();
            chooser.setMultiSelectionEnabled(true);
            FileNameExtensionFilter filter = new FileNameExtensionFilter(
                    "Spreadsheets", "xls", "xlsx");
                chooser.setFileFilter(filter);
                int returnVal = chooser.showOpenDialog(null);
                if(returnVal == JFileChooser.APPROVE_OPTION) {
                    files = chooser.getSelectedFiles();
                } else {
                    System.out.println("No file chosen!");
                    return;
                }
        } else {
           files = new File[] { new File(args[0]) } ;
        }
        for (File f : files) {
            FileInputStream fis = new FileInputStream(f);
            Workbook wb = null;
            if (f.getName().endsWith(".xlsx")) {
                wb = new XSSFWorkbook(fis);
            } else {
                wb = new HSSFWorkbook(fis);
            }
            List<SheetReader> sheets = new ArrayList<SheetReader>();
            for (int i = 0; i < wb.getNumberOfSheets(); i ++) {
                sheets.add(new SheetReader(wb.getSheetAt(i)));
            }

            File outputFile = new File(f.getName().substring(0, f.getName().lastIndexOf(".") + 1).concat("txt"));
            PrintWriter writer = new PrintWriter(new OutputStreamWriter(new FileOutputStream(outputFile), "ISO-8859-1"));

            while (true) {
                int lowest = sheets.get(0).getLowestNumber();
                SheetReader lowestSheet = sheets.get(0);
    
                for (int i = 1; i < sheets.size(); i ++) {
                    SheetReader s = sheets.get(i);
                    int l = s.getLowestNumber();
                    if (l != -1 && (lowest == -1 || l < lowest)) {
                        lowest = l;
                        lowestSheet = s;
                    }
                }
                if (lowest == -1) {
                    break;
                } else {
                    lowestSheet.serializeRow(writer);
                }
            }
            writer.close();
            JOptionPane.showMessageDialog(null, "Created PLUM TXT file: " + outputFile.getAbsolutePath());
        }
    }

    public static void convertSpreadsheet(File inputFile) throws IOException {
        FileInputStream fis = new FileInputStream(inputFile);
        Workbook wb = null;
        if (inputFile.getName().endsWith(".xlsx")) {
            wb = new XSSFWorkbook(fis);
        } else {
            wb = new HSSFWorkbook(fis);
        }
        List<SheetReader> sheets = new ArrayList<SheetReader>();
        for (int i = 0; i < wb.getNumberOfSheets(); i ++) {
            sheets.add(new SheetReader(wb.getSheetAt(i)));
        }

        String outputFilename = inputFile.getName().substring(0, inputFile.getName().lastIndexOf(".") + 1).concat("txt");
        while (new File(outputFilename).exists()) {
            DateFormat f = new SimpleDateFormat("yyyy-MMM-dd.HH.mm.ss");
            outputFilename = outputFilename.replace(".txt", "-" + f.format(new Date()) + ".txt");
        }
        PrintWriter writer = new PrintWriter(new OutputStreamWriter(new FileOutputStream(outputFilename), "ISO-8859-1"));

        int count = 0;
        while (true) {
            int lowest = sheets.get(0).getLowestNumber();
            SheetReader lowestSheet = sheets.get(0);
            for (int i = 1; i < sheets.size(); i ++) {
                SheetReader s = sheets.get(i);
                int l = s.getLowestNumber();
                if (l != -1 && (lowest == -1 || l < lowest)) {
                    lowest = l;
                    lowestSheet = s;
                }
            }
            if (lowest == -1) {
                break;
            } else {
                count ++;
                lowestSheet.serializeRow(writer);
            }
        }
        writer.close();
        JOptionPane.showMessageDialog(null, "Created PLUM TXT file (" + count + " records): " + new File(outputFilename).getAbsolutePath());

    }
    
    public static final class SheetReader {
        private Row header;

        private Row next;

        private Iterator<Row> rows;

        public SheetReader(Sheet s) {
            rows = s.rowIterator();
            header = rows.next();
            next = null;
        }
        private Row getCurrentRow() {
            if (next == null) {
                if (rows.hasNext()) {
                    next = rows.next();
                } else {
                    next = null;
                }
            }
            return next;
        }

        private void advanceRow() {
            next = null;
        }

        public int getLowestNumber() {
            if (getCurrentRow() == null) {
                return -1;
            }
            return Integer.parseInt(getCurrentRow().getCell(0).getStringCellValue());
        }

        public void serializeRow(PrintWriter w) {
            // The first line doesn't appear to end i the FIELD_DELIMITER in the
            // example files I saw, so we will omit it from this
            boolean firstLine = getLowestNumber() == 0;
            StringBuffer line = new StringBuffer();
            for (int i = 1; i < header.getLastCellNum(); i ++) {
                line.append(header.getCell(i).getStringCellValue());
                Cell c = getCurrentRow().getCell(i);
                line.append(c == null ? "" : c.getStringCellValue());
                line.append(FIELD_DELIMITER);
            }
            System.out.println(line);
            advanceRow();
            w.println(firstLine ? line.substring(0, line.length() - 1) : line);
        }
    }
}
