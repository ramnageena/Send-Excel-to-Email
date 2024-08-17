package com.excel.repository;

import com.excel.entity.PDOData;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface PDORepository extends JpaRepository<PDOData,String> {
}
