/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package filemerger;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.logging.Level;
import java.util.logging.Logger;
import static javax.script.ScriptEngine.FILENAME;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author mmercado
 */
public class Excel {

    public static XSSFWorkbook getSummaryFile() {
        String[] headers = {"ID Number", "Last Name", "First Name", "Position",
            "Month", "Year", "Department", "Product", "Date Submitted",
            "Approving Manager", "Date Approved", "Week Day", "Saturday",
            "Sunday/Rest Day/Overtime", "Special NonWorking Holiday",
            "Legal Holiday", "Shift Allowance", "Night Differential", "Meal",
            "Transportation", "Others", "Count Meal",
            "Count Meal/Transportation", "Count Shift Allowance",
            "Count OT Days", "Count Day", "", "", "Form Version", "", "ID",
            "OT Weekday", "OT Sunday", "OT SH", "OT LH", "ND", "Count SA", "SA",
            "Count M+T", "M+T"};

        String[] edHeaders = {"ID Number", "Last Name", "First Name", "Position",
            "Month Year", "Department", "Product",
            "Date Submitted", "Approving Manager",
            "Date Approved", "Count Day", "", "", "Form Version",
            "", "ID Number", "Last Name", "First Name",
            "Position", "Month", "Year", "Department", "Product",
            "Count Day"};

        XSSFWorkbook workbook = new XSSFWorkbook();

        workbook.createSheet("OT");
        workbook.createSheet("SA");
        workbook.createSheet("AHS");
        workbook.createSheet("ED");
        for (int i = 0; i < 4; i++) {
            XSSFSheet sheet = workbook.getSheetAt(i);
            //if sheet is not ED sheet
            if (i < 3) {
                Row row = sheet.createRow(0);
                for (int j = 0; j < headers.length; j++) {
                    Cell cell = row.createCell(j);
                    cell.setCellValue(headers[j]);
                    sheet.setColumnWidth(j, 6000);
                }
            } else {
                Row row = sheet.createRow(0);
                for (int j = 0; j < edHeaders.length; j++) {
                    Cell cell = row.createCell(j);
                    cell.setCellValue(edHeaders[j]);
                    sheet.setColumnWidth(j, 6000);
                }
            }

        }
        return workbook;

    }

    public static ArrayList<Form> getClaimsData(Model model, View view) {
        ArrayList<Form> forms = new ArrayList();
        int progress = 1;
        view.getProgressBar().setMaximum(model.getFiles().length);
        for (File file : model.getFiles()) {
            view.getProgressBar().setValue(progress);
            view.getProgressBar().update(view.getProgressBar().getGraphics());
            try {
                Form form = new Form();
                String id = file.getName();
                id = id.substring(0, id.length() - 5);
                id = id.substring(id.length() - 5);
                form.setType(file.getName().substring(0, 3).trim());
                form.setId(id);
                Workbook workbook = WorkbookFactory.create(file, model.getPasswords().get(id));
                Sheet firstSheet = workbook.getSheetAt(0);
                if (firstSheet.getProtect()) {
                    form.setHack("ok");
                } else {
                    form.setHack("tampered");
                }
                setPersonalCells(form, workbook);
                for (int rowIndex = 18; rowIndex < 49; rowIndex++) {
                    Row row = firstSheet.getRow(rowIndex);
                    switch (row.getCell(1).getCellTypeEnum()) {
                        case STRING:
                            if (!row.getCell(1).getStringCellValue().equals("")) {
                                form.setCountDay(form.getCountDay() + 1);
                            }
                            break;
                        case NUMERIC:
                            if (row.getCell(1).getNumericCellValue() > 0) {
                                form.setCountDay(form.getCountDay() + 1);
                            }
                            break;
                    }

                    for (int cellIndex = 9; cellIndex < 21; cellIndex++) {
                        setNumericCells(cellIndex, row.getCell(cellIndex), form);
                    }
                }
                Row row = firstSheet.getRow(51);
                String formVersion = row.getCell(1).getStringCellValue();
                form.setFormVersion(formVersion.substring(formVersion.length() - 10));
                if(!row.getCell(21).getStringCellValue().equals("")){
                    form.setGenerated("AUTO");
                }else{
                    form.setGenerated("MANUAL");
                }
                forms.add(form);
                workbook.close();
                progress++;
            } catch (EncryptedDocumentException ex) {
                model.getErrors().add(file.getName() + " supplied password is invalid");
                progress++;

            } catch (IOException ex) {
                model.getErrors().add(file.getName() + " has been set to read only");
                progress++;

            } catch (InvalidFormatException ex) {
                Logger.getLogger(Excel.class
                        .getName()).log(Level.SEVERE, null, ex);
                progress++;
            }
        }
        return forms;
    }

    public static void createSummaryFile(Workbook workbook, ArrayList<Form> forms, View view) throws FileNotFoundException, IOException {
        DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HHmmss");
        Date date = new Date();
        String filename = dateFormat.format(date);
        FileOutputStream fileOutputStream = new FileOutputStream(filename + ".xlsx");
        HashMap<String, Integer> rowNums = new HashMap();
        initRowMap(rowNums);
        view.getProgressBar().setValue(0);
        view.getProgressBar().update(view.getProgressBar().getGraphics());
        view.getProgressBar().setMinimum(0);
        view.getProgressBar().setMaximum(forms.size());
        int progress = 1;
        for (Form form : forms) {
            view.getProgressBar().setValue(progress);
            view.getProgressBar().update(view.getProgressBar().getGraphics());
            Sheet sheet = workbook.getSheet(form.getType());
            Row row = sheet.createRow(rowNums.get(form.getType()));
            int rowNum = rowNums.get(form.getType());
            int cellNum = 0;
            if (form.getType().equals("ED")) {
                row.createCell(cellNum++).setCellValue(form.getId());
                row.createCell(cellNum++).setCellValue(form.getLastName());
                row.createCell(cellNum++).setCellValue(form.getFirstName());
                row.createCell(cellNum++).setCellValue(form.getPosition());
                row.createCell(cellNum++).setCellValue(form.getMonth());
                row.createCell(cellNum++).setCellValue(form.getYear());
                row.createCell(cellNum++).setCellValue(form.getDepartment());
                row.createCell(cellNum++).setCellValue(form.getProduct());
                row.createCell(cellNum++).setCellValue(form.getDateSubmitted());
                row.createCell(cellNum++).setCellValue(form.getApprovingManager());
                row.createCell(cellNum++).setCellValue(form.getDateApproved());
                row.createCell(cellNum++).setCellValue(form.getCountDay());
                row.createCell(cellNum++).setCellValue(form.getGenerated());
                row.createCell(cellNum++).setCellValue(form.getHack());
                row.createCell(cellNum++).setCellValue(form.getFormVersion());
                row.createCell(cellNum++).setCellFormula("A" + (rowNum + 1));
                row.createCell(cellNum++).setCellFormula("B" + (rowNum + 1));
                row.createCell(cellNum++).setCellFormula("C" + (rowNum + 1));
                row.createCell(cellNum++).setCellFormula("D" + (rowNum + 1));
                row.createCell(cellNum++).setCellFormula("E" + (rowNum + 1));
                row.createCell(cellNum++).setCellFormula("F" + (rowNum + 1));
                row.createCell(cellNum++).setCellFormula("G" + (rowNum + 1));
                row.createCell(cellNum++).setCellFormula("H" + (rowNum + 1));
                row.createCell(cellNum++).setCellFormula("L" + (rowNum + 1));
                rowNums.put(form.getType(), ++rowNum);
                progress++;
            } else {
                row.createCell(cellNum++).setCellValue(form.getId());
                row.createCell(cellNum++).setCellValue(form.getLastName());
                row.createCell(cellNum++).setCellValue(form.getFirstName());
                row.createCell(cellNum++).setCellValue(form.getPosition());
                row.createCell(cellNum++).setCellValue(form.getMonth());
                row.createCell(cellNum++).setCellValue(form.getYear());
                row.createCell(cellNum++).setCellValue(form.getDepartment());
                row.createCell(cellNum++).setCellValue(form.getProduct());
                row.createCell(cellNum++).setCellValue(form.getDateSubmitted());
                row.createCell(cellNum++).setCellValue(form.getApprovingManager());
                row.createCell(cellNum++).setCellValue(form.getDateApproved());
                row.createCell(cellNum++).setCellValue(form.getWeekDay());
                row.createCell(cellNum++).setCellValue(form.getSaturday());
                row.createCell(cellNum++).setCellValue(form.getRestDay());
                row.createCell(cellNum++).setCellValue(form.getNoneWorkingDay());
                row.createCell(cellNum++).setCellValue(form.getLegalHoliday());
                row.createCell(cellNum++).setCellValue(form.getShiftAllowance());
                row.createCell(cellNum++).setCellValue(form.getNightDifferential());
                row.createCell(cellNum++).setCellValue(form.getMeal());
                row.createCell(cellNum++).setCellValue(form.getTransportation());
                row.createCell(cellNum++).setCellValue(form.getOthers());
                row.createCell(cellNum++).setCellValue(form.getCountMeal());
                row.createCell(cellNum++).setCellValue(form.getCountTransportation());
                row.createCell(cellNum++).setCellValue(form.getCountShiftAllowance());
                row.createCell(cellNum++).setCellValue(form.getCountOTDays());
                row.createCell(cellNum++).setCellValue(form.getCountDay());
                row.createCell(cellNum++).setCellValue(form.getGenerated());
                row.createCell(cellNum++).setCellValue(form.getHack());
                row.createCell(cellNum++).setCellValue(form.getFormVersion());
                cellNum++;
                row.createCell(cellNum++).setCellFormula("A" + (rowNum + 1));
                row.createCell(cellNum++).setCellFormula("SUM(L" + (rowNum + 1) + ":M" + (rowNum + 1) + ")");
                row.createCell(cellNum++).setCellFormula("SUM(N" + (rowNum + 1) + ")");
                row.createCell(cellNum++).setCellFormula("SUM(O" + (rowNum + 1) + ")");
                row.createCell(cellNum++).setCellFormula("SUM(P" + (rowNum + 1) + ")");
                row.createCell(cellNum++).setCellFormula("R" + (rowNum + 1));
                row.createCell(cellNum++).setCellFormula("X" + (rowNum + 1));
                row.createCell(cellNum++).setCellFormula("Q" + (rowNum + 1));
                row.createCell(cellNum++).setCellFormula("SUM(V" + (rowNum + 1) + ":W" + (rowNum + 1) + ")");
                row.createCell(cellNum++).setCellFormula("SUM(S" + (rowNum + 1) + ":T" + (rowNum + 1) + ")");
                rowNums.put(form.getType(), ++rowNum);
                progress++;
            }
        }
        workbook.write(fileOutputStream);
        workbook.close();
    }

    //function to get personal information
    public static void setPersonalCells(Form form, Workbook workbook) {
        Sheet firstSheet = workbook.getSheetAt(0);
        Row row = firstSheet.getRow(9);
        form.setLastName(row.getCell(5).getStringCellValue());
        form.setFirstName(row.getCell(8).getStringCellValue());
        form.setPosition(row.getCell(11).getStringCellValue());
        form.setMonth(row.getCell(14).getStringCellValue());
        form.setYear((int) row.getCell(17).getNumericCellValue());
        row = firstSheet.getRow(12);
        form.setDepartment(row.getCell(2).getStringCellValue());
        form.setProduct(row.getCell(5).getStringCellValue());
        setDateCells(row.getCell(8), form, 8);
        if (row.getCell(11).getStringCellValue() != null) {
            form.setApprovingManager(row.getCell(11).getStringCellValue());
        } else {
            form.setApprovingManager("");
        }
        setDateCells(row.getCell(17), form, 17);
    }

    public static void setDateCells(Cell cell, Form form, int index) {
        DateFormat df = new SimpleDateFormat("MM-dd-yyyy");
        switch (cell.getCellTypeEnum()) {
            case STRING:
                if (cell.getStringCellValue() != null) {
                    if (index == 8) {
                        form.setDateSubmitted(cell.getStringCellValue());
                    } else {
                        form.setDateApproved(cell.getStringCellValue());
                    }

                } else {
                    if (index == 8) {
                        form.setDateSubmitted("");
                    } else {
                        form.setDateApproved("");
                    }

                }
                break;
            case NUMERIC:
                if (cell.getDateCellValue() != null) {
                    if (index == 8) {
                        form.setDateSubmitted(df.format(cell.getDateCellValue()));
                    } else {
                        form.setDateApproved(df.format(cell.getDateCellValue()));
                    }

                } else {
                    if (index == 8) {
                        form.setDateSubmitted("");
                    } else {
                        form.setDateApproved("");
                    }

                }
                break;
        }
    }

    public static void setNumericCells(int index, Cell cell, Form form) {
        switch (index) {
            case 9:
                form.setWeekDay(form.getWeekDay() + cell.getNumericCellValue());
                if (cell.getNumericCellValue() > 0 && !form.isCheckOTDays()) {
                    form.setCountOTDays(form.getCountOTDays() + 1);
                    form.setCheckOTDays(true);
                }
                break;
            case 10:
                form.setSaturday(form.getSaturday() + cell.getNumericCellValue());
                if (cell.getNumericCellValue() > 0 && !form.isCheckOTDays()) {
                    form.setCountOTDays(form.getCountOTDays() + 1);
                    form.setCheckOTDays(true);
                }
                break;
            case 11:
                form.setRestDay(form.getRestDay() + cell.getNumericCellValue());
                if (cell.getNumericCellValue() > 0 && !form.isCheckOTDays()) {
                    form.setCountOTDays(form.getCountOTDays() + 1);
                    form.setCheckOTDays(true);
                }
                break;
            case 12:
                form.setNoneWorkingDay(form.getNoneWorkingDay() + cell.getNumericCellValue());
                if (cell.getNumericCellValue() > 0 && !form.isCheckOTDays()) {
                    form.setCountOTDays(form.getCountOTDays() + 1);
                    form.setCheckOTDays(true);
                }
                break;
            case 13:
                form.setLegalHoliday(form.getLegalHoliday() + cell.getNumericCellValue());
                if (cell.getNumericCellValue() > 0 && !form.isCheckOTDays()) {
                    form.setCountOTDays(form.getCountOTDays() + 1);
                    form.setCheckOTDays(true);
                }
                break;
            case 15:
                switch (cell.getCachedFormulaResultTypeEnum()) {
                    case ERROR:
                        form.setShiftAllowance(form.getShiftAllowance() + 0);
                        break;
                    case NUMERIC:
                        form.setShiftAllowance(form.getShiftAllowance() + cell.getNumericCellValue());
                        if (cell.getNumericCellValue() > 0 && !form.isCheckShiftAllowance()) {
                            form.setCountShiftAllowance(form.getCountShiftAllowance() + 1);
                            form.setCheckShiftAllowance(true);
                        }
                        break;
                }
                break;
            case 16:
                switch (cell.getCachedFormulaResultTypeEnum()) {
                    case ERROR:
                        form.setNightDifferential(form.getNightDifferential() + 0);
                        break;
                    case NUMERIC:
                        form.setNightDifferential(form.getNightDifferential() + cell.getNumericCellValue());
                        break;
                }
                break;
            case 17:
                switch (cell.getCachedFormulaResultTypeEnum()) {
                    case ERROR:
                        form.setMeal(form.getMeal() + 0);
                        break;
                    case NUMERIC:
                        form.setMeal(form.getMeal() + cell.getNumericCellValue());
                        if (cell.getNumericCellValue() > 0 && !form.isCheckMeal()) {
                            form.setCountMeal(form.getCountMeal() + 1);
                            form.setCheckMeal(true);
                        }
                        break;
                }
                break;
            case 18:
                switch (cell.getCachedFormulaResultTypeEnum()) {
                    case ERROR:
                        form.setTransportation(form.getTransportation() + 0);
                        break;
                    case NUMERIC:
                        form.setTransportation(form.getTransportation() + cell.getNumericCellValue());
                        if (cell.getNumericCellValue() > 0 && form.isCheckMeal() && !form.isCheckTransportation()) {
                            form.setCountTransportation(form.getCountTransportation() + 1);
                            form.setCheckTransportation(true);
                        }
                        break;
                }
                break;
            case 19:
                switch (cell.getCellTypeEnum()) {
                    case NUMERIC:
                        form.setOthers(form.getOthers() + cell.getNumericCellValue());
                        break;
                    case STRING:
                        break;
                }
                break;
            case 20:
                form.setCheckDay(false);
                form.setCheckMeal(false);
                form.setCheckOTDays(false);
                form.setCheckTransportation(false);
                form.setCheckShiftAllowance(false);
                break;
        }
    }

    public static void initRowMap(HashMap<String, Integer> rowNums) {
        rowNums.put("OT", 1);
        rowNums.put("SA", 1);
        rowNums.put("AHS", 1);
        rowNums.put("ED", 1);
    }

    public static void writeErrorToFile(ArrayList<String> errors) throws IOException {
        BufferedWriter bw = null;
        FileWriter fw = null;

        try {

            DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HHmmss");
            Date date = new Date();
            String timestamp = dateFormat.format(date);
            File file = new File("ErrorLog" + timestamp  + ".txt");

            // if file doesnt exists, then create it
            if (!file.exists()) {
                file.createNewFile();
            }

            // true = append file
            fw = new FileWriter(file.getAbsoluteFile(), true);
            bw = new BufferedWriter(fw);
            for (String error : errors) {
                  bw.write(error);
                  bw.newLine();
            }
          

            System.out.println("Done");

        } catch (IOException e) {
        } finally {

            if (bw != null) {
                bw.close();
            }
            if (fw != null) {
                fw.close();
            }
        }
    }
}
