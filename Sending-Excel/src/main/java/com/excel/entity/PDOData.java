package com.excel.entity;

import jakarta.persistence.*;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.math.BigDecimal;
import java.sql.Date;
@Entity
@Table(name = "PDO_Data")
@Data
@AllArgsConstructor
@NoArgsConstructor
public class PDOData {
    @Id

    private String ccy;
    private BigDecimal amount;
    private Date maturityDate;
    private String referenceNumber;
}
