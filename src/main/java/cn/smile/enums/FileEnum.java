package cn.smile.enums;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.Getter;

@AllArgsConstructor
@Getter
public enum FileEnum {


    EXCEL("excel", ".xls"),
    EXCEL_HIGH("excel", ".xlsx"),
    ;
    private String name;
    private String value;
}
