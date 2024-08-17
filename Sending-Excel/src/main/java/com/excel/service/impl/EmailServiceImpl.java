package com.excel.service.impl;

import com.excel.exception.EmailSendingException;
import com.excel.service.EmailService;
import com.excel.service.ExcelGenerate;
import jakarta.mail.MessagingException;
import jakarta.mail.internet.MimeMessage;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.mail.javamail.MimeMessageHelper;
import org.springframework.stereotype.Service;

import java.io.ByteArrayInputStream;
import java.io.IOException;

@Service
@Slf4j
public class EmailServiceImpl implements EmailService {
    @Autowired
    private JavaMailSender mailSender;

    @Autowired
    private ExcelGenerate excelGenerate;


    @Override
    public void sendEmailWithExcel() throws MessagingException, IOException {
        log.info("Starting email preparation with Excel attachment");

        try {
            ByteArrayInputStream excelStream = excelGenerate.createExcel();
            byte[] byteArray = readExcelStream(excelStream);
            ByteArrayResource attachment = new ByteArrayResource(byteArray);

            MimeMessage message = mailSender.createMimeMessage();
            MimeMessageHelper helper = createMimeMessageHelper(message, attachment);

            mailSender.send(message);
            log.info("Email with Excel attachment sent successfully");

        } catch (MessagingException | IOException e) {
            log.error("Error occurred while sending email with Excel attachment", e);
            throw new EmailSendingException("Failed to send email with Excel attachment", e);
        }
    }

    private byte[] readExcelStream(ByteArrayInputStream excelStream) throws IOException {
        return excelStream.readAllBytes();
    }

    private MimeMessageHelper createMimeMessageHelper(MimeMessage message, ByteArrayResource attachment) throws MessagingException {
        MimeMessageHelper helper = new MimeMessageHelper(message, true);
        helper.setTo("ramtiwari8716@gmail.com");
        helper.setSubject("Data Excel Report");
        helper.setText("Please find the attached PDO Data report.");
        helper.addAttachment("PDOData.xlsx", attachment);
        return helper;
    }


}



















//    ByteArrayInputStream excelStream =excelGenerate.createExcel();
//
//        // Read all the bytes into a byte array
//        byte[] byteArray = excelStream.readAllBytes();
//
//        // Create a ByteArrayResource from the byte array
//        ByteArrayResource attachment = new ByteArrayResource(byteArray);
//
//        MimeMessage message = mailSender.createMimeMessage();
//        MimeMessageHelper helper = new MimeMessageHelper(message, true);
//
//        helper.setTo("ramtiwari8716@gmail.com");
//        helper.setSubject("Data Excel Report");
//        helper.setText("Please find the attached for PDO Data report.");
//
//        helper.addAttachment("PDOData.xlsx", attachment);
//
//        mailSender.send(message);
 //   }
//}
