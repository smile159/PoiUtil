package cn.smile.util;//import lombok.extern.slf4j.Slf4j;

import cn.smile.annotation.Excel;
import cn.smile.entity.User;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import javax.xml.ws.Response;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;

@Slf4j
public class PoiExcelUtil {


    // 工作簿对象
    private final Workbook workbook;


    // 样式表
    private Map<String, CellStyle> styles;


    public PoiExcelUtil() {
        this.workbook = new SXSSFWorkbook();
        this.init();
    }


    public void init() {
        styles = new HashMap<>();
        // 创建样式对象
        CellStyle publicCellStyle = workbook.createCellStyle();
        // 设置水平对齐方式：居中
        publicCellStyle.setAlignment(HorizontalAlignment.CENTER);
        // 设置垂直对齐方式：剧中
        publicCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        Font publicFont = workbook.createFont();
        publicFont.setFontName("Arial");
        publicFont.setFontHeightInPoints((short) 10);
        publicCellStyle.setFont(publicFont);
        styles.put("public", publicCellStyle);

        CellStyle headCellStyle = workbook.createCellStyle();
        // 复制公共样式表
        headCellStyle.cloneStyleFrom(publicCellStyle);
        // 设置背景色
        headCellStyle.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        headCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font headFont = workbook.createFont();
        headFont.setFontName("Arial");
        headFont.setFontHeightInPoints((short) 10);
        headFont.setColor(IndexedColors.WHITE.getIndex());
        headCellStyle.setFont(headFont);
        styles.put("head", headCellStyle);
        log.info("PoiExcelUtil init done");
    }


    private void writeHeader(Sheet sheet, String[] headList, CellStyle cellStyle) {
        // 创建第一行
        SXSSFRow row = (SXSSFRow) sheet.createRow(0);
        for (int i = 0; i < headList.length; i++) {
            String cellValue = headList[i];
            SXSSFCell cell = row.createCell(i);
            cell.setCellValue(cellValue);
            cell.setCellStyle(cellStyle);
        }
        log.info("PoiExcelUtil writeHeader done");
    }


    private void writeHeader(Sheet sheet, List<String> headList, CellStyle cellStyle) {
        // 创建第一行
        SXSSFRow row = (SXSSFRow) sheet.createRow(0);
        for (int i = 0; i < headList.size(); i++) {
            String cellValue = headList.get(i);
            SXSSFCell cell = row.createCell(i);
            cell.setCellValue(cellValue);
            cell.setCellStyle(cellStyle);
        }
        log.info("PoiExcelUtil writeHeader done");
    }


    private <T> void writeHeader(Sheet sheet, T t, CellStyle cellStyle) {
        SXSSFRow row = (SXSSFRow) sheet.createRow(0);
        Class<?> tClass = t.getClass();
        Field[] declaredFields = tClass.getDeclaredFields();
        int count = 0;
        for (Field field : declaredFields) {
            Annotation[] annotations = field.getAnnotations();
            for (Annotation annotation : annotations) {
                if (annotation.annotationType() == Excel.class) {
                    SXSSFCell cell = row.createCell(count++);
                    Excel excel = (Excel) annotation;
                    String value = excel.value();
                    cell.setCellValue(value);
                    cell.setCellStyle(cellStyle);
                }
            }
        }
        log.info("PoiExcelUtil writeHeader done");
    }


    private void writeData(Sheet sheet, List<List<String>> dataList, CellStyle cellStyle) {
        for (int i = 0; i < dataList.size(); i++) {
            SXSSFRow row = (SXSSFRow) sheet.createRow(i + 1);
            // 获取列数据
            List<String> columnList = dataList.get(i);
            for (int j = 0; j < columnList.size(); j++) {
                String cellValue = columnList.get(j);
                SXSSFCell cell = row.createCell(j);
                cell.setCellValue(cellValue);
                cell.setCellStyle(cellStyle);
            }
        }
        log.info("PoiExcelUtil writeData done");
    }

    private <T> void writeDtoData(Sheet sheet, List<T> dataList, CellStyle cellStyle) {
        for (int i = 0; i < dataList.size(); i++) {
            SXSSFRow row = (SXSSFRow) sheet.createRow(i + 1);
            T t = dataList.get(i);
            Class<?> dataClass = t.getClass();
            Field[] dataDeclaredFields = dataClass.getDeclaredFields();
            int count = 0;
            for (Field field : dataDeclaredFields) {
                field.setAccessible(true);
                Annotation[] annotations = field.getAnnotations();
                for (Annotation annotation : annotations) {
                    if (annotation.annotationType() == Excel.class) {
                        try {
                            SXSSFCell cell = row.createCell(count++);
                            // 获取字段类型
                            Class<?> fieldType = field.getType();
                            // 获取字段具体值
                            Object fieldValue = field.get(t);
//                            System.out.println(fieldType.getTypeName() + ":" + field.getName() + ":" + fieldValue);
                            // 暂时全部当作String写入excel，以后做处理
                            if (fieldValue != null) {
                                cell.setCellValue(fieldValue.toString());
                                cell.setCellStyle(cellStyle);
                            }
                        } catch (IllegalAccessException e) {
                            e.printStackTrace();
                        }
                    }
                }
            }
        }
        log.info("PoiExcelUtil writeData done");
    }


    private void checkParams(List<String> headList, List<List<String>> dataList, OutputStream outputStream) {
        if (headList.isEmpty() || dataList.isEmpty()) {
            log.error("【参数错误】：缺少参数headList");
            throw new RuntimeException("【参数错误】：缺少参数headList");
        }
        if (outputStream == null) {
            log.error("【参数错误】：缺少参数outputStream");
            throw new RuntimeException("【参数错误】：缺少参数outputStream");
        }
    }

    private void checkParams(List<String> headList, List<List<String>> dataList, HttpServletResponse response) {
        if (headList.isEmpty() || dataList.isEmpty()) {
            log.error("【参数错误】：缺少参数headList");
            throw new RuntimeException("【参数错误】：缺少参数headList");
        }
        if (response == null) {
            log.error("【参数错误】：缺少参数response");
            throw new RuntimeException("【参数错误】：缺少参数response");
        }
    }


    private void writeAndCloseWorkbook(OutputStream outputStream) {
        try {
            workbook.write(outputStream);
            log.info("PoiExcelUtil write Excel done");
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            if (workbook != null) {
                try {
                    workbook.close();
                    log.info("PoiExcelUtil exportExcel close");
                } catch (IOException e) {
                    log.error("PoiExcelUtil exportExcel close error");
                    throw new RuntimeException(e);
                }
            }
        }
    }


    public void exportExcel(List<String> headList, List<List<String>> dataList, OutputStream outputStream) throws IOException {
        // 检查参数
        checkParams(headList, dataList, outputStream);
        exportExcel(headList, "public", dataList, "public", outputStream);
    }

    public void exportExcel(List<String> headList, List<List<String>> dataList, String fileName, String fileSuffix, HttpServletResponse response) throws IOException {
        // 检查参数
        checkParams(headList, dataList, response);
        // 这里还需要对response添加响应头
        ResponseUtil.setFileResponseHeader(response, fileName + fileSuffix);
        exportExcel(headList, "head", dataList, "public", response.getOutputStream());
    }

    private void exportExcel(List<String> headList, String headStyle, List<List<String>> dataList, String dataStyle, OutputStream outputStream) throws IOException {
        Sheet sheet = workbook.createSheet();
        writeHeader(sheet, headList, styles.get(headStyle));
        writeData(sheet, dataList, styles.get(dataStyle));
        writeAndCloseWorkbook(outputStream);
    }

    public <T> void exportExcel(List<T> dataList, OutputStream outputStream) {
        Sheet sheet = workbook.createSheet();
        // 写入表头
        this.writeHeader(sheet, dataList.get(0), styles.get("head"));
        // 写入数据
        this.writeDtoData(sheet, dataList, styles.get("public"));
        writeAndCloseWorkbook(outputStream);
    }


    public static void main(String[] args) throws IOException {
        PoiExcelUtil poiExcelUtil = new PoiExcelUtil();
        OutputStream outputStream = Files.newOutputStream(Paths.get("D://baseExcel.xlsx"));


//        List<String> headList = Arrays.asList("工号", "姓名", "年龄", "性别", "组织部门");
//        List<List<String>> dataList = Arrays.asList(
//                Arrays.asList("1", "张三", "23", "男", "财务部"),
//                Arrays.asList("2", "李四", "24", "男", "财务部"),
//                Arrays.asList("3", "王五", "25", "男", "财务部")
//        );
//        poiExcelUtil.exportExcel(headList, dataList, outputStream);


        List<User> userList = new ArrayList<>();
        userList.add(new User(null, "张三", 0, "a", "123@qq.com", "123456789", "1", null, null));
        userList.add(new User(null, "李四", 0, "a", "123@qq.com", "123456789", "1", null, null));
        userList.add(new User(null, "王五", 0, "a", "123@qq.com", "123456789", "1", null, null));
        poiExcelUtil.exportExcel(userList, outputStream);


    }


}
