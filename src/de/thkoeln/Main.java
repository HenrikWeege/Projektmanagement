package de.thkoeln;



import org.apache.poi.ss.usermodel.*;

import javax.swing.JFileChooser;
import java.io.File;

public class Main {

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

            // Optimize machine planning and scheduling
            myBasicOptimizer.process();

            // Write optimized data back to excel sheet
            writeExcelFileData();

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

    private static void readExcelFileData(String fileName){
        try {
            Workbook workbook = WorkbookFactory.create(new File(fileName));

            // this is an example, modify code here
            Sheet sheet = workbook.getSheetAt(0);
            Row row = sheet.getRow(1);
            Cell cell = row.getCell(0);

            System.out.println("Value read (A/2): " + cell.getStringCellValue());
        } catch (Exception exReadExcelFile) {
            System.out.println("Error in readExcelFileData()");
            System.out.println("Error details: " + exReadExcelFile.toString());
        }
    }

    private static void writeExcelFileData(){
        try {
            System.out.println("Called writeExcelFileData()");
        } catch (Exception exWriteExcelFile) {
            System.out.println("Error in writeExcelFileData()");
            System.out.println("Error details: " + exWriteExcelFile.toString());
        }
    }
}
Test123
