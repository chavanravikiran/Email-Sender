package com.emailSender;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.io.IOException;

public class EmployeeSearch {
	public static void main(String[] args) {
        String employeeId = "2050"; // value to search

        boolean result = searchEmployee(employeeId);
        if (result) {
            System.out.println("Employee ID " + employeeId + " found in Excel.");
        } else {
            System.out.println("Employee ID " + employeeId + " not found.");
        }
    }

    public static boolean searchEmployee(String employeeId) {
        boolean found = false;
        String filePath = "src/main/resources/employees.xlsx"; // path to your file

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // first sheet

            for (Row row : sheet) {
                for (Cell cell : row) {
                    cell.setCellType(CellType.STRING); // convert everything to String
                    String value = cell.getStringCellValue().trim();

                    if (value.equals(employeeId)) {
                        found = true;
                        break;
                    }
                }
                if (found) break; // break outer loop if already found
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

        return found;
    }
}
