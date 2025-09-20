package com.emailSender.service;

import java.io.File;
import java.io.FileInputStream;
import java.util.Arrays;
import javax.mail.internet.MimeMessage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.FileSystemResource;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.mail.javamail.MimeMessageHelper;
import org.springframework.stereotype.Service;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;

@Service
public class BulkEmailService {

	@Autowired
	private JavaMailSender mailSender;
	
	@Value("$spring.mail.username")
	private String userName;
	
	
	
	public void sendEmailsFromExcel(String excelPath, String resumePath,String currentCompanyName,String yearExperience,String skils,String resumeName,String coverPagePath,String emailTitle) {
		try (FileInputStream fis = new FileInputStream(excelPath); Workbook workbook = new XSSFWorkbook(fis)) {

			Sheet sheet = workbook.getSheetAt(0); // First sheet

			for (int i = 1; i <= sheet.getLastRowNum(); i++) { // Start from 1 to skip header
				Row row = sheet.getRow(i);
				if (row == null) {
					System.out.println("Row " + (i + 1) + " is empty or null. Skipping.");
					continue;
				}
				
				String companyName = getCellValue(row.getCell(0)); // Company
				String applylocation = getCellValue(row.getCell(1));//Location
				String emailField = getCellValue(row.getCell(3)); // Email

				if (emailField == null || emailField.trim().isEmpty()) {
			        System.out.println("Skipping row " + (i + 1) + ": Email is empty.");
			        continue;
			    }
				
				System.out.println("Processing Row " + (i + 1) + ": Sending email to " + emailField + " at " + companyName);
				String[] emailIds = emailField.split("\\r?\\n"); // for multiple IDs

				 try {
//		                sendEmail(emailIds, companyName, resumePath, currentCompanyName, applylocation, yearExperience, skils,resumeName);
		                sendEmailWithCoverPageAndResume(emailIds,resumePath,resumeName,coverPagePath,emailTitle);
		                System.out.println("✅ Email sent successfully for row " + (i + 1));
		            } catch (Exception ex) {
		                System.out.println("❌ Failed to send email for row " + (i + 1) + ": " + ex.getMessage());
		            }
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	

	private void sendEmail(String[] recipients, String companyName, String resumePath,String currentCompanyName,String applylocation,String yearExperience,String skils,String resumeName) {
//		try {
//			MimeMessage message = mailSender.createMimeMessage();
//			MimeMessageHelper helper = new MimeMessageHelper(message, true);
//
//			helper.setTo(recipients);
//			helper.setSubject("Inquiry: Java / Spring Boot Opportunity at " + companyName);
//
//			String emailContent ="Hi,\n\n"
//					+"I hope you're doing well. I came across the Java Backend Developer role at "+companyName+" in "+applylocation+" and wanted to express my interest.\r\n"
//					+ "with "+yearExperience+" years of experience in "+skils+", I believe I could be a strong fit.\r\n\n"
//					+ "Could you please let me know the right channel to apply or connect with the relevant hiring team?\r\n\n"
//					+ "Looking forward to your response.\r\n\n"
//					+ "Best Regards,\n" + "Ravikiran Chavan \n"
//					+ "https://www.linkedin.com/in/ravikiran-chavan-2846b1188/\r\n"
//					+ "https://ravikiran-portfolio.vercel.app/ ";
//			
//			helper.setText(emailContent);
//
//			// Attach Resume
//			FileSystemResource file = new FileSystemResource(new File(resumePath));
//			helper.addAttachment(resumeName, file);
//
//			mailSender.send(message);
//			System.out.println("Email sent to: " + Arrays.toString(recipients));
//
//		} catch (Exception e) {
//			e.printStackTrace();
//		}
	}
	

	private String getCellValue(Cell cell) {
		if (cell == null)
			return "";
		if (cell.getCellType() == CellType.STRING)
			return cell.getStringCellValue().trim();
		else if (cell.getCellType() == CellType.NUMERIC)
			return String.valueOf((int) cell.getNumericCellValue());
		return "";
	}

	private void sendEmailWithCoverPageAndResume(String[] recipients, String resumePath, String resumeName,
			String coverPagePath,String emailTitle) {
		
		try {
			MimeMessage message = mailSender.createMimeMessage();
			MimeMessageHelper helper = new MimeMessageHelper(message, true);
			helper.setFrom(new InternetAddress(userName, emailTitle));

			helper.setTo(recipients);
			helper.setSubject("Application for Java Developer");

			String emailContent =
	                "<p>Greetings,</p>" +
	                "<p>I hope you're doing well. I’m currently looking for opportunities as a " +
	                "<b>Java/Spring Boot Backend Developer</b> and came across your profile.</p>" +

	                "<p>With <b>3+ years of experience</b> in Java, Spring Boot, REST APIs, " +
	                "Microservices, PostgreSQL, Oracle Database and related technologies, " +
	                "I’m confident in my ability to contribute to dynamic backend teams.</p>" +

	                "<p>I’ve attached my resume and cover page for your reference. " +
	                "If there are any relevant openings at your organization, or if you could refer me, " +
	                "I’d be truly grateful.</p>" +

	                "<p>Thank you for your time and consideration!</p>" +
//	                "<br>" +
	                "<p>Regards,<br>" +
	                "<b>Ravikiran Chavan</b><br>" +
	                "📞 +91 8552805879<br>" +
	                "<a href=\"https://www.linkedin.com/in/your-profile\">LinkedIn</a> | " +
	                "<a href=\"https://your-portfolio.com\">Portfolio</a>" +
	                "</p>";
	

			helper.setText(emailContent, true);
			FileSystemResource file = new FileSystemResource(new File(resumePath));
	        helper.addAttachment(resumeName, file);

	        // Attach Cover Page (optional)
	        if (coverPagePath != null && !coverPagePath.isEmpty()) {
	            FileSystemResource coverPage = new FileSystemResource(new File(coverPagePath));
	            helper.addAttachment("Ravikiran_Chavan_CoverLetter.pdf", coverPage);
	        }

	        mailSender.send(message);
	        System.out.println("Email sent to: " + Arrays.toString(recipients));
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
