package durbin.mike.plum;

import java.io.File;
import java.io.IOException;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;

public class BidirectionalConverter {

    public static void main(String [] args) throws IOException {
        File[] files = null;
        if (args.length == 0) {
            JFileChooser chooser = new JFileChooser();
            chooser.setMultiSelectionEnabled(true);
            FileNameExtensionFilter filter = new FileNameExtensionFilter(
                    "Plum Spreadsheets or Text Files", "xls", "xlsx", "txt");
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
            if (f.getName().endsWith(".txt")) {
                PlumTXTToSpreadsheet.convertText(f);
            } else {
                PlumSpreadsheetToTXT.convertSpreadsheet(f);
            }
        }
    }
}
