package com.excel.service;

import jakarta.mail.MessagingException;

import java.io.IOException;

public interface EmailService {
    public void sendEmailWithExcel() throws MessagingException, IOException;
}
