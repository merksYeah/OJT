package filemerger;

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
/**
 *
 * @author mmercado
 */
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.JProgressBar;
import javax.swing.text.BadLocationException;
import javax.swing.text.StyledDocument;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Controller {

    private View view;
    private Model model;

    public Controller(Model m, View v) {
        model = m;
        view = v;
        initView();
    }

    public void initView() {
        view.getSelectPane().setEditable(false);
        view.getBtnExtract().setEnabled(false);
        view.getScrollPane().setViewportView(view.getSelectPane());
    }

    public void initController() {
        view.getBtnSelect().addActionListener(e -> getFiles());
        view.getBtnExtract().addActionListener(e -> extractFiles());
    }

    private void getFiles() {
        view.getSelectPane().setText("");
        File passFile = new File("Passwords.xlsx");
        if (passFile.exists()) {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setMultiSelectionEnabled(true);
            int result = fileChooser.showOpenDialog(fileChooser);
            File[] files = fileChooser.getSelectedFiles();
            if (result == JFileChooser.APPROVE_OPTION) {
                if (files.length > 0) {
                    //Gets Passwords and Puts them in Model Hashmap
                    try {
                        FileInputStream excelFile = new FileInputStream(passFile);
                        Workbook workbook = new XSSFWorkbook(excelFile);
                        Sheet datatypeSheet = workbook.getSheetAt(0);
                        Iterator<Row> rowIterator = datatypeSheet.iterator();
                        DataFormatter formatter = new DataFormatter();
                        int index = 0;
                        while (rowIterator.hasNext()) {
                            Row row = rowIterator.next();
                            if (index > 0) {
                                String key = formatter.formatCellValue(row.getCell(0));
                                String value = formatter.formatCellValue(row.getCell(1));
                                model.getPasswords().put(key, value);
                            }
                            index++;
                        }
                        //Appends file names to text pane
                        for (File file : files) {
                            StyledDocument doc = view.getSelectPane().getStyledDocument();
                            doc.insertString(doc.getLength(), file.getName() + "\n", null);
                        }
                    } catch (IOException | BadLocationException ex) {
                        Logger.getLogger(Controller.class.getName()).log(Level.SEVERE, null, ex);
                    }
                    model.setFiles(files);
                    view.getBtnExtract().setEnabled(true);
                    JOptionPane.showMessageDialog(view.getFrame(),
                            "You can now proceed to extracting",
                            "Successfully selected files",
                            JOptionPane.INFORMATION_MESSAGE);
                }
            }
        } else {
            JOptionPane.showMessageDialog(view.getFrame(),
                    "Passwords.xlsx cannot be found",
                    "File cannot be found",
                    JOptionPane.ERROR_MESSAGE);
        }

    }

    private void extractFiles() {
        try {
            XSSFWorkbook summaryWorkbook = Excel.getSummaryFile();
            model.setForms(Excel.getClaimsData(model, view));
            if (model.getErrors().size() > 0) {
                Excel.writeErrorToFile(model.getErrors());
                JOptionPane.showMessageDialog(view.getFrame(),
                        "Please see generated error log for details",
                        "Errors Encountered",
                        JOptionPane.ERROR_MESSAGE);
            } else {
                JOptionPane.showMessageDialog(view.getFrame(),
                        "All data was extracted press ok to generate summary file",
                        "Data Extraction Completed",
                        JOptionPane.INFORMATION_MESSAGE);
            }
            Excel.createSummaryFile(summaryWorkbook, model.getForms(), view);
            JOptionPane.showMessageDialog(view.getFrame(),
                    "Summary File has been generated",
                    "Processing Done",
                    JOptionPane.INFORMATION_MESSAGE);
            view.getProgressBar().setValue(0);
            view.getProgressBar().update(view.getProgressBar().getGraphics());
            view.getBtnExtract().setEnabled(false);
            view.getSelectPane().setText("");
            model.getErrors().clear();
        } catch (IOException ex) {
            Logger.getLogger(Controller.class.getName()).log(Level.SEVERE, null, ex);
        }

    }
}
