package com.excel.service.impl;

import com.excel.entity.PDOData;
import com.excel.exception.ExcelGenerationException;
import com.excel.repository.PDORepository;
import com.excel.service.ExcelGenerate;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.sql.Date;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import static com.excel.constant.QueryConstant.*;

@Service
@Slf4j
public class ExcelGenerateImpl implements ExcelGenerate {

    @Autowired
    private JdbcTemplate jdbcTemplate;

    @Autowired
    private PDORepository pdoRepository;


    @Override
    public ByteArrayInputStream createExcelwithoutProReady() throws IOException {
        List<Map<String, Object>> data = jdbcTemplate.queryForList(SQL_QUERY);
        Set<String> uniqueReferenceNumber = new HashSet<>();

        String SUMMERY = "SUMMERY";

        try (Workbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            Sheet sheet = workbook.createSheet("Data");
            // Create Header

            /* Font started */
            CellStyle headerCellStyle = workbook.createCellStyle();
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            headerCellStyle.setFont(headerFont);
            /* Font ended */


            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < HEADERS.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(HEADERS[i]);
                cell.setCellStyle(headerCellStyle); //Font
            }
            int rowNum = 1;
            for (Map<String, Object> row : data) {
                Row dataRow = sheet.createRow(rowNum++);
                for (int i = 0; i < HEADERS.length; i++) {
                    dataRow.createCell(i).setCellValue(row.get(HEADERS[i]) != null ? row.get(HEADERS[i]).toString() : "");
                }
                uniqueReferenceNumber.add(row.getOrDefault("ReferenceNumber", " ").toString());
                dataRow.createCell(HEADERS.length - 1).setCellValue("DummyData");
            }
            Sheet summerySheet = workbook.createSheet(SUMMERY);

            Row summeryRow = summerySheet.createRow(0);
            Cell sumCel = summeryRow.createCell(0);
            sumCel.setCellValue(SUMMERY);
            sumCel.setCellStyle(headerCellStyle);

            int sumNum = 1;
            for (String s : uniqueReferenceNumber) {
                Row row = summerySheet.createRow(sumNum++);
                row.createCell(0).setCellValue(s);
            }
            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        }
    }

    @Override
    public ByteArrayInputStream createExcel() throws IOException {
        log.info("Starting Excel generation process");
        List<Map<String, Object>> dataFromDb;
        Set<String> uniqueReferenceNumber = new HashSet<>();
        String SUMMARY = "SUMMARY";

        try {
            dataFromDb = fetchDataFromDb();
            log.debug("Data fetched from database: {}", dataFromDb);
            saveDataToPdoTable(dataFromDb);
            log.info("Data store in PDO table successfully");

        } catch (Exception e) {
            log.error("Error fetching data from database", e);
            throw new ExcelGenerationException("Error fetching data from database", e);
        }

        try (Workbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream out = new ByteArrayOutputStream()) {

            createDataSheet(workbook, dataFromDb, uniqueReferenceNumber);
            createSummarySheet(workbook, uniqueReferenceNumber, SUMMARY);

            workbook.write(out);
            log.info("Excel generation completed successfully");
            return new ByteArrayInputStream(out.toByteArray());

        } catch (IOException e) {
            log.error("Error while writing Excel file", e);
            throw new ExcelGenerationException("Error while writing Excel file", e);
        }


    }

    private List<Map<String, Object>> fetchDataFromDb() {
        List<Map<String, Object>> dataFromDb = jdbcTemplate.queryForList(SQL_QUERY);
        return dataFromDb;
    }

    private void createDataSheet(Workbook workbook, List<Map<String, Object>> data, Set<String> uniqueReferenceNumber) {
        Sheet sheet = workbook.createSheet(SHEET_NAME1);
        CellStyle headerCellStyle = createHeaderCellStyle(workbook);

        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < HEADERS.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(HEADERS[i]);
            cell.setCellStyle(headerCellStyle);
        }

        int rowNum = 1;
        for (Map<String, Object> row : data) {
            Row dataRow = sheet.createRow(rowNum++);
            for (int i = 0; i < HEADERS.length; i++) {
                dataRow.createCell(i).setCellValue(row.get(HEADERS[i]) != null ? row.get(HEADERS[i]).toString() : "");
            }
            uniqueReferenceNumber.add(row.getOrDefault("referenceNumber", " ").toString());
            dataRow.createCell(HEADERS.length - 1).setCellValue("DummyData");
        }
    }

    private void createSummarySheet(Workbook workbook, Set<String> uniqueReferenceNumber, String sheetName) {
        Sheet summarySheet = workbook.createSheet(sheetName);
        CellStyle headerCellStyle = createHeaderCellStyle(workbook);

        Row summaryRow = summarySheet.createRow(0);
        Cell summaryCell = summaryRow.createCell(0);
        summaryCell.setCellValue(sheetName);
        summaryCell.setCellStyle(headerCellStyle);

        int rowNum = 1;
        for (String refNumber : uniqueReferenceNumber) {
            Row row = summarySheet.createRow(rowNum++);
            row.createCell(0).setCellValue(refNumber);
        }
    }

    private CellStyle createHeaderCellStyle(Workbook workbook) {
        CellStyle headerCellStyle = workbook.createCellStyle();
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerCellStyle.setFont(headerFont);
        return headerCellStyle;
    }

    private void saveDataToPdoTable(List<Map<String, Object>> data) {
        for (Map<String, Object> mapData : data) {
            PDOData pdoData = new PDOData();
            pdoData.setCcy(mapData.get("ccy").toString());
            pdoData.setAmount(new BigDecimal(mapData.get("amount").toString()));
            pdoData.setMaturityDate(Date.valueOf(mapData.get("maturityDate").toString()));
            pdoData.setReferenceNumber(mapData.get("referenceNumber").toString());
            pdoRepository.save(pdoData);
        }

    }
}
