package com.excel.controller;

import com.excel.service.EmailService;
import com.excel.service.ExcelGenerate;
import jakarta.mail.MessagingException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.ByteArrayInputStream;
import java.io.IOException;

@RestController
@RequestMapping("/api")
public class SendEmailController {

    @Autowired
    ExcelGenerate excelGenerate;

    @Autowired
    EmailService emailService;

    @GetMapping("/download")
    public ResponseEntity<byte[]> downloadExcel() throws IOException {
        ByteArrayInputStream byteArrayInputStream = excelGenerate.createExcel();

        HttpHeaders headers = new HttpHeaders();
        headers.add("Content-Disposition", "attachment; filename=PDOData.xlsx");

        return ResponseEntity.ok()
                .headers(headers)
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .body(byteArrayInputStream.readAllBytes());
    }

    @GetMapping("/sendEmail")
    public String sendEmail() {
        try {
            emailService.sendEmailWithExcel();
            return "Email sent successfully!";
        } catch (MessagingException | IOException e) {
            return "Failed to send email: " + e.getMessage();
        }
    }

}
