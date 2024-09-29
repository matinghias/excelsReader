package com.excelsreader;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@SpringBootApplication
public class ExcelsReaderApplication {

    public static void main(String[] args) throws IOException {
        String file1Path = "excel1.xlsx";  // .xlsx file (target)
        String file2Path = "excel2.xls";   // .xls file (source)

        // Load Excel files
        FileInputStream file1InputStream = new FileInputStream(file1Path);
        Workbook workbook1;
        if (file1Path.endsWith(".xls")) {
            workbook1 = new HSSFWorkbook(file1InputStream);  // For .xls files
        } else {
            workbook1 = new XSSFWorkbook(file1InputStream);  // For .xlsx files
        }

        FileInputStream file2InputStream = new FileInputStream(file2Path);
        Workbook workbook2;
        if (file2Path.endsWith(".xls")) {
            workbook2 = new HSSFWorkbook(file2InputStream);  // For .xls files
        } else {
            workbook2 = new XSSFWorkbook(file2InputStream);  // For .xlsx files
        }

        // Assuming we're working with the first sheets in both files
        Sheet sheet1 = workbook1.getSheetAt(0);  // Target sheet (Excel 1)
        Sheet sheet2 = workbook2.getSheetAt(0);  // Source sheet (Excel 2)

        // Map to store (Name -> H value) from file2 (source)
        Map<String, List<String>> nameToHValuesMap = new HashMap<>();

        // Read from file2: column L (index 11) and column H (index 7)
        for (int i = 1; i <= sheet2.getLastRowNum(); i++) {
            Row rowFile2 = sheet2.getRow(i);
            if (rowFile2 != null) {
                Cell nameCell = rowFile2.getCell(11);  // Column L (names) in file2
                Cell hValueCell = rowFile2.getCell(7);  // Column H (values) in file2

                if (nameCell != null && hValueCell != null) {
                    // Store name and corresponding H value in the map
                    String name = nameCell.toString().trim();
                    String hValue = hValueCell.toString().trim();

                    // If name already exists, add the value to the list, otherwise create a new list
                    nameToHValuesMap.computeIfAbsent(name, k -> new ArrayList<>()).add(hValue);
                }
            }
        }

        // Iterate over file1: column F (index 5) and update column J (index 9) if name matches
        for (int i = 1; i <= sheet1.getLastRowNum(); i++) {
            Row rowFile1 = sheet1.getRow(i);
            if (rowFile1 != null) {
                Cell nameCellInFile1 = rowFile1.getCell(5);  // Column F in file1 (target)

                if (nameCellInFile1 != null) {
                    String nameInFile1 = nameCellInFile1.toString().trim();

                    // Check if the name exists in the map
                    if (nameToHValuesMap.containsKey(nameInFile1)) {
                        // Retrieve the list of H values corresponding to the name
                        List<String> hValues = nameToHValuesMap.get(nameInFile1);

                        // If the list has values, we will assign the first value and remove it from the list
                        if (!hValues.isEmpty()) {
                            // Write the corresponding H value into column J (index 9) of file1
                            Cell jColumnCell = rowFile1.createCell(9);  // Column J
                            jColumnCell.setCellValue(hValues.remove(0));  // Remove the first element from the list
                        }
                    }
                }
            }
        }

        // Close the input streams before writing to the file
        file1InputStream.close();
        file2InputStream.close();

        // Write the changes to file1
        FileOutputStream file1OutputStream = new FileOutputStream(file1Path);
        workbook1.write(file1OutputStream);

        // Close all resources
        workbook1.close();
        workbook2.close();
        file1OutputStream.close();

        System.out.println("Excel comparison and update completed!");
    }

}
