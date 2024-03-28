package cn.smile.controller;

import cn.smile.enums.FileEnum;
import cn.smile.util.PoiExcelUtil;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

@RestController
@RequestMapping("excel")
public class ExcelController {




    @GetMapping("download")
    public void downloadUserExcel(HttpServletResponse response){
        PoiExcelUtil poiExcelUtil = new PoiExcelUtil();
        List<String> headList = Arrays.asList("工号", "姓名", "年龄", "性别", "组织部门");
        List<List<String>> dataList = Arrays.asList(
                Arrays.asList("1", "张三", "23", "男", "财务部"),
                Arrays.asList("2", "李四", "24", "男", "财务部"),
                Arrays.asList("3", "王五", "25", "男", "财务部")
        );
        try {
            poiExcelUtil.exportExcel(headList, dataList, "测试Excel文件名", FileEnum.EXCEL_HIGH.getValue(), response);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
