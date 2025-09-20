package com.emailSender.runner;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.CommandLineRunner;
import org.springframework.stereotype.Component;
import com.emailSender.service.BulkEmailService;

@Component
public class AppRunner implements CommandLineRunner {
	
	@Autowired
    private BulkEmailService bulkEmailService;

    public void run(String... args) {
        String excelPath = "E:/Study Materials-Spring Boot/Resume/emails.xlsx";
        String resumePath = "E:/Study Materials-Spring Boot/Resume/RavikiranChavan_Backend_Dev.pdf";
        String coverPagePath = "E:/Study Materials-Spring Boot/Resume/Ravikiran_Chavan_Cover_Page.pdf";
        String currentCompanyName = "Probity Software Pvt Ltd.";
        String yearExperience = "3+";
        String skils = "Java, Spring Boot, Spring MVC, Spring Data, REST APIs, Microservices, PostgreSQL, Oracle, and Jasper Reports";
        String resumeName ="Ravikiran_Chavan_Resume.pdf";
        String emailTitle ="Ravikiran Chavan";
        
        bulkEmailService.sendEmailsFromExcel(excelPath, resumePath,currentCompanyName,yearExperience, skils,resumeName,coverPagePath,emailTitle);
    }
}
