package de.thkoeln;



import org.apache.poi.ss.usermodel.*;

import javax.swing.*;
import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

public class Main {
    static Workbook workbook;
    static Sheet sheet;

    public static void main(String[] args) {
        try {
            // Variables
            Optimizer myBasicOptimizer = new Optimizer();
            System.out.println("Optimizer instantiated...");

            // Select Excel File which the user wants to investigate regarding machine optimization potential
            String excelFileName = selectExcelFile();
            System.out.println("File via openFileDialog selected: " + excelFileName);

            // Read out data from excel sheet via Apache POI
            readExcelFileData(excelFileName);

            List<Row> rows = sortExcelFileData(19);
            // Optimize machine planning and scheduling
            myBasicOptimizer.process();

            // Write optimized data back to excel sheet
            writeExcelFileData("C:\\Users\\Henrik\\Desktop\\Neuer Ordner\\ProductionSheet NEU.xlsx", rows, 1);

            System.out.println("Optimization done.");
        } catch (Exception ex) {
            System.out.println(ex.toString());
        }
    }

    // Select Excel file which we want to read out and retrieve necessary machine data
    private static String selectExcelFile(){
        String strFilename = "";

        try {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setCurrentDirectory(new File(System.getProperty("user.home")));

            int result = fileChooser.showOpenDialog(null);

            switch(result) {
                case JFileChooser.APPROVE_OPTION:
                    File selectedFile = fileChooser.getSelectedFile();

                    // user made a selection -> return selected file path
                    strFilename = selectedFile.toString();
                    break;

                case JFileChooser.CANCEL_OPTION:
                    // user made no selection -> shutdown application
                    System.out.println("No selection made. Shutting down application...");
                    System.exit(0);
                    break;

                default:
                    System.out.println("Something went wrong. Shutting down application...");
                    System.exit(0);
                    break;
            }
        } catch (Exception exFileSelection) {
            System.out.println("Error in selectExcelFile()");
            System.out.println("Error details: " + exFileSelection.toString());
        } finally {
            return strFilename;
        }
    }

    private static List<Row> sortExcelFileData(int sortColumn)
    {
        List<Row> rows = new ArrayList<>();
        try {

            for (int i = 1; i<sheet.getPhysicalNumberOfRows();i ++ )
            {
                rows.add(sheet.getRow(i)); //(new SortRow(sheet.getRow(i).getCell(19).getStringCellValue(), sheet.getRow(i)));
            }
            // ToDo: Sortieren bei Zahlenwerten? Keine Ahnung Finn/Hauke
            rows.sort((r1, r2) -> r1.getCell(sortColumn).getStringCellValue().compareTo(r2.getCell(sortColumn).getStringCellValue()));

        } catch (Exception exp) {
            System.out.println("Error in writeExcelFileData()");
            System.out.println("Error details: " + exp.toString());
        }
        return rows;
    }

    private static void readExcelFileData(String fileName){
        try {
            workbook = WorkbookFactory.create(new File(fileName));
            sheet = workbook.getSheetAt(1);
        } catch (Exception exReadExcelFile) {
            System.out.println("Error in readExcelFileData()");
            System.out.println("Error details: " + exReadExcelFile.toString());
        }
    }

    private static void writeExcelFileData(String filename, List<Row> rows, int startRow){
        try {
            for (int i = 0; i < rows.size(); i++)
            {
                Row row = sheet.getRow(i + startRow);
                for (int x = 0; x < rows.get(i).getPhysicalNumberOfCells(); x++)
                {
                    if (rows.get(i).getCell(x).getCellType() == CellType.STRING) {
                        row.createCell(x).setCellValue(rows.get(i).getCell(x).getStringCellValue());
                        //Fall:  Zelle ist ein Buchstabe; Erstellt identische Zelle in identischer Zeile und weißt dieser Zelle ihren entsprechenden Wert zu, dieser Wert wird daraufhin in Buchstaben umgewandelt
                    } else if (rows.get(i).getCell(x).getCellType() == CellType.NUMERIC) {
                        row.createCell(x).setCellValue(rows.get(i).getCell(x).getNumericCellValue());
                        //Fall:  Zelle ist ein Nummererischer Wert; Erstellt identische Zelle in identischer Zeile und weißt dieser Zelle ihren entsprechenden Wert zu, dieser Wert wird daraufhin in Zahl umgewandelt
                    } else if (rows.get(i).getCell(x).getCellType() == CellType.FORMULA) {
                        String cellFormula = rows.get(i).getCell(x).getCellFormula();
                        System.out.println(cellFormula);
                        row.createCell(x).setCellFormula(cellFormula);
                        //Fall:  Zelle ist eine Formel; Erstellt identische Zelle in identischer Zeile und weißt dieser Zelle ihren entsprechenden Wert zu, dieser Wert wird daraufhin in Formelwert umgewandelt
                    } else {
                        System.out.println(rows.get(i).getCell(x).getCellType());
                    }

                }
            }
            try (FileOutputStream outputStream = new FileOutputStream(filename)) {
                workbook.write(outputStream);
            }
            System.out.println("Called writeExcelFileData()");
        } catch (Exception exWriteExcelFile) {
            System.out.println("Error in writeExcelFileData()");
            System.out.println("Error details: " + exWriteExcelFile.toString());
        }
    }
}

