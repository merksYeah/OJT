/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package filemerger;

import java.io.File;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
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
            "Month", "Year", "Department", "Product",
            "Date Submitted", "Approving Manager",
            "Date Approved", "Count Day", "", "",
            "Form Version", "", "ID Number", "Last Name", "First Name",
            "Position	Month", "Year", "Department", "Product",
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
                    cell.setCellValue(headers[j]);
                    sheet.setColumnWidth(j, 6000);
                }
            }

        }
        return workbook;

    }

    public static ArrayList<Form> getClaimsData(HashMap<String, String> passwords, File[] files) {
        ArrayList<Form> forms = new ArrayList();
        for (File file : files) {
            try {
                Form form = new Form();
                String id = file.getName().substring(file.getName().length() - 10, file.getName().length() - 5);
                form.setType(file.getName().substring(0, 3));
                form.setId(id);
                Workbook workbook = WorkbookFactory.create(file, passwords.get(id));
                Sheet firstSheet = workbook.getSheetAt(0);
                setPersonalCells(form, workbook);
                for (int i = 18; i < 49; i++) {
                    Row row = firstSheet.getRow(i);
                    System.out.println(row.getCell(9).getNumericCellValue());
                    /*for(int j = 9; j < 21; j++){
                        
                    }*/
                }

            } catch (IOException | InvalidFormatException | EncryptedDocumentException ex) {
                System.out.println("hello");
            }
        }
        return forms;
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
        setDateCells(row.getCell(8), form);
        if (row.getCell(11).getStringCellValue() != null) {
            form.setApprovingManager(row.getCell(11).getStringCellValue());
        } else {
            form.setApprovingManager("");
        }
        setDateCells(row.getCell(17), form);
    }

    public static void setDateCells(Cell cell, Form form) {
        DateFormat df = new SimpleDateFormat("MM-dd-yyyy");
        switch (cell.getCellTypeEnum()) {
            case STRING:
                if (cell.getStringCellValue() != null) {
                    form.setDateSubmitted(cell.getStringCellValue());
                } else {
                    form.setDateSubmitted("");
                }
                break;
            case NUMERIC:
                if (cell.getDateCellValue() != null) {
                    form.setDateSubmitted(df.format(cell.getDateCellValue()));
                } else {
                    form.setDateSubmitted("");
                }
                break;
        }
    }
    
    public static void setNumericCells(int index){
        switch(index){
            case 9:
                
        }
    }

}
