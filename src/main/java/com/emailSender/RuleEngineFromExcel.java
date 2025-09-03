package com.emailSender;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import javax.script.ScriptEngine;
import javax.script.ScriptEngineManager;
import javax.script.SimpleBindings;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

class Applicant {
    int age;
    double salary;
    String country;

    Applicant(int age, double salary, String country) {
        this.age = age;
        this.salary = salary;
        this.country = country;
    }

    public void approveApplication() {
        System.out.println("✅ Application Approved");
    }

    public void rejectApplication() {
        System.out.println("❌ Application Rejected");
    }

    public void applyRegionalBenefits() {
        System.out.println("🌍 Regional Benefits Applied");
    }
}

public class RuleEngineFromExcel {

    public static void main(String[] args) {
        // ✅ Separate files for applicants and rules
        String applicantsFile = "D:\\studyPurpose\\JavaExamples\\src\\originaldatafile.xlsx"; 
        String rulesFile = "D:\\studyPurpose\\JavaExamples\\src\\rules.xlsx"; 
        String resultsFile = "D:\\studyPurpose\\JavaExamples\\src\\results.xlsx"; 

        try {
            // Load applicants from Excel
            List<Applicant> applicants = loadApplicants(applicantsFile);

            // Load rules from Excel
            List<String[]> rules = loadRules(rulesFile);

            // Nashorn engine (Java 1.8)
            ScriptEngine engine = new ScriptEngineManager().getEngineByName("nashorn");
            if (engine == null) {
                throw new RuntimeException("❌ Nashorn engine not available. Please use Java 1.8 or 11.");
            }

            // Create results workbook
            Workbook resultsWorkbook = new XSSFWorkbook();
            Sheet resultsSheet = resultsWorkbook.createSheet("Results");

            // Header row
            Row header = resultsSheet.createRow(0);
            header.createCell(0).setCellValue("Applicant");
            header.createCell(1).setCellValue("Age");
            header.createCell(2).setCellValue("Salary");
            header.createCell(3).setCellValue("Country");
            header.createCell(4).setCellValue("Condition");
            header.createCell(5).setCellValue("Action");
            header.createCell(6).setCellValue("Result");

            int resultRowIndex = 1;

            // Process each applicant
            int applicantId = 1;
            for (Applicant applicant : applicants) {
                Map<String, Object> variables = new HashMap<>();
                variables.put("applicant", applicant);

                for (String[] rule : rules) {
                    String condition = rule[0];
                    String action = rule[1];

                    boolean result = (boolean) engine.eval(condition, new SimpleBindings(variables));

                    String executionResult = "SKIPPED";
                    if (result) {
                        String methodName = action.replace("();", "").trim();
                        try {
                            Method method = Applicant.class.getMethod(methodName);
                            method.invoke(applicant);
                            executionResult = "EXECUTED";
                        } catch (Exception e) {
                            executionResult = "FAILED: " + e.getMessage();
                        }
                    }

                    // Write result row
                    Row resultRow = resultsSheet.createRow(resultRowIndex++);
                    resultRow.createCell(0).setCellValue("Applicant_" + applicantId);
                    resultRow.createCell(1).setCellValue(applicant.age);
                    resultRow.createCell(2).setCellValue(applicant.salary);
                    resultRow.createCell(3).setCellValue(applicant.country);
                    resultRow.createCell(4).setCellValue(condition);
                    resultRow.createCell(5).setCellValue(action);
                    resultRow.createCell(6).setCellValue(executionResult);
                }
                applicantId++;
            }

            // Save results file
            try (FileOutputStream fos = new FileOutputStream(resultsFile)) {
                resultsWorkbook.write(fos);
            }
            resultsWorkbook.close();

            System.out.println("✅ Results saved in " + resultsFile);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // Load applicants from Excel file
    private static List<Applicant> loadApplicants(String applicantsFile) throws IOException {
        List<Applicant> applicants = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(new File(applicantsFile));
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            int rowCount = sheet.getPhysicalNumberOfRows();
            DataFormatter formatter = new DataFormatter();

            for (int i = 1; i < rowCount; i++) { // start at row 1, skip headers
                Row row = sheet.getRow(i);
                if (row == null) continue;

                String ageStr = formatter.formatCellValue(row.getCell(0));
                String salaryStr = formatter.formatCellValue(row.getCell(1));
                String country = formatter.formatCellValue(row.getCell(2));

                if (ageStr.isEmpty() || salaryStr.isEmpty() || country.isEmpty()) {
                    continue; // skip blank rows
                }

                try {
                    int age = Integer.parseInt(ageStr.trim());
                    double salary = Double.parseDouble(salaryStr.trim());
                    applicants.add(new Applicant(age, salary, country));
                } catch (NumberFormatException e) {
                    System.out.println("⚠️ Skipping invalid row: " + ageStr + ", " + salaryStr + ", " + country);
                }
            }
        }
        return applicants;
    }

    // Load rules from Excel file
    private static List<String[]> loadRules(String rulesFile) throws IOException {
        List<String[]> rules = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(new File(rulesFile));
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            int rowCount = sheet.getPhysicalNumberOfRows();
            DataFormatter formatter = new DataFormatter();

            for (int i = 1; i < rowCount; i++) { // start at row 1, skip headers
                Row row = sheet.getRow(i);
                if (row == null) continue;

                String condition = formatter.formatCellValue(row.getCell(0));
                String action = formatter.formatCellValue(row.getCell(1));
                if (!condition.isEmpty() && !action.isEmpty()) {
                    rules.add(new String[]{condition, action});
                }
            }
        }
        return rules;
    }
}
