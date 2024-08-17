package com.excel.constant;

public class QueryConstant {
    public static final String[] HEADERS = {"ReferenceNumber", "MaturityDate", "CCY", "Amount", "ABC"};
    public static final String SQL_QUERY = "Select db.ccy,db.amount,db.maturityDate,db.referenceNumber from excel_db db LIMIT 10; ";
    public static final String SHEET_NAME1 = "PDO Data";
}
