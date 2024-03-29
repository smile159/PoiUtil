package cn.smile.entity;

import cn.smile.annotation.Excel;
import lombok.AllArgsConstructor;
import lombok.Data;
import org.apache.poi.ss.usermodel.CellType;
import org.springframework.beans.factory.annotation.Value;

import java.util.Date;

@Data
@AllArgsConstructor
public class Test {


    @Excel("int")
    private int t1;

    @Excel("Integer")
    private Integer t2;

    @Excel("String")
    private String t3;

    @Excel(value = "Double", cellType = CellType.NUMERIC, dateFormat = "0.00")
    private Double t4;

    @Excel(value = "Float", dateFormat = "0.0")
    private Float t5;

    @Excel("Long")
    private Long t6;

    @Excel("Short")
    private Short t7;

    @Excel("Byte")
    private Byte t8;

    @Excel("Boolean")
    private Boolean t9;

    @Excel("Character")
    private Character t10;

    @Excel(value = "Date", dateFormat = "yyyy-MM-dd")
    private Date t11;

    @Excel(value = "java.sql.Date", dateFormat = "yyyy-MM-dd")
    private java.sql.Date t12;

    @Excel(value = "java.sql.Time", dateFormat = "hh:mm:ss")
    private java.sql.Time t13;

    @Excel(value = "java.sql.Timestamp", dateFormat = "yyyy-MM-dd hh:mm:ss")
    private java.sql.Timestamp t14;

    @Excel(value = "java.time.LocalDate", dateFormat = "yyyy-MM-dd")
    private java.time.LocalDate t17;

    @Excel("java.time.LocalTime")
    private java.time.LocalTime t19;

    @Excel(value = "java.time.LocalDateTime", dateFormat = "yyyy-MM-dd hh:mm:ss")
    private java.time.LocalDateTime t18;


}
