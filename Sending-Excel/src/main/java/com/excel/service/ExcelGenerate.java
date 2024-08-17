package com.excel.service;


import java.io.ByteArrayInputStream;
import java.io.IOException;

public interface ExcelGenerate {

    ByteArrayInputStream createExcel() throws IOException;
    ByteArrayInputStream createExcelwithoutProReady() throws IOException;
}
