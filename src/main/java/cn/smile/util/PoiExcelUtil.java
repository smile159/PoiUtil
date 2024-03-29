package cn.smile.util;

import cn.hutool.core.util.IdUtil;
import cn.smile.annotation.Excel;
import cn.smile.entity.Test;
import cn.smile.entity.User;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.sql.Time;
import java.sql.Timestamp;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.*;
import java.util.stream.Collectors;

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


    public List<Field> filterNoExcelField(Field[] fields) {
        List<Field> fieldList = new ArrayList<>();
        for (Field field : fields) {
            Annotation[] annotations = field.getAnnotations();
            for (Annotation annotation : annotations) {
                if (annotation.annotationType() == Excel.class) {
                    fieldList.add(field);
                }
            }
        }
        return fieldList.stream().sorted(Comparator.comparingInt(o -> o.getAnnotation(Excel.class).sort())).collect(Collectors.toList());
    }


    private <T> void writeHeader(Sheet sheet, T t, CellStyle cellStyle) {
        SXSSFRow row = (SXSSFRow) sheet.createRow(0);
        Class<?> tClass = t.getClass();
        List<Field> fields = filterNoExcelField(tClass.getDeclaredFields());
        for (int i = 0; i < fields.size(); i++) {
            Field field = fields.get(i);
            Annotation[] annotations = field.getAnnotations();
            for (Annotation annotation : annotations) {
                if (annotation.annotationType() == Excel.class) {
                    SXSSFCell cell = row.createCell(i);
                    Excel excel = (Excel) annotation;
                    // 设置列的宽度
                    sheet.setColumnWidth(i, excel.width() * 256);
                    // 设置行的高度
                    row.setHeightInPoints(excel.height());
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


    public void parseDataType(Field field, Object fieldValue, Excel excel, SXSSFCell cell, CellStyle cellStyle) {
        CellStyle dataCellStyle = workbook.createCellStyle();
        dataCellStyle.cloneStyleFrom(cellStyle);
        Class<?> fieldType = field.getType();
        cell.setCellType(excel.cellType());
        // 根据不同类型执行不同操作
        if (fieldType == int.class || fieldType == Integer.class) {
            System.out.println("这是一个整数类型（int 或 Integer）");
            cell.setCellValue((Integer) fieldValue);
        } else if (fieldType == String.class) {
            System.out.println("这是一个字符串类型（String）");
            cell.setCellValue((String) fieldValue);
        } else if (fieldType == double.class || fieldType == Double.class) {
            System.out.println("这是一个双精度浮点数类型（double 或 Double）");
            cell.setCellValue((Double) fieldValue);
        } else if (fieldType == float.class || fieldType == Float.class) {
            System.out.println("这是一个单精度浮点数类型（float 或 Float）");
            cell.setCellValue((Float) fieldValue);
        } else if (fieldType == Date.class) {
            System.out.println("这是一个日期类型（Date）");
            cell.setCellValue((Date) fieldValue);
        } else if (fieldType == Boolean.class) {
            System.out.println("这是一个布尔类型（Boolean）");
            cell.setCellValue((Boolean) fieldValue);
        } else if (fieldType == byte.class || fieldType == Byte.class) {
            System.out.println("这是一个字节类型（byte 或 Byte）");
            cell.setCellValue((Byte) fieldValue);
        } else if (fieldType == char.class || fieldType == Character.class) {
            System.out.println("这是一个字符类型（char 或 Character）");
            cell.setCellValue((Character) fieldValue);
        } else if (fieldType == long.class || fieldType == Long.class) {
            System.out.println("这是一个长整数类型（long 或 Long）");
            cell.setCellValue((Long) fieldValue);
        } else if (fieldType == short.class || fieldType == Short.class) {
            System.out.println("这是一个短整数类型（short 或 Short）");
        } else if (fieldType == BigDecimal.class) {
            cell.setCellValue(((BigDecimal) fieldValue).doubleValue());
            System.out.println("这是一个高精度浮点数类型（BigDecimal）");
        } else if (fieldType == LocalDateTime.class) {
            System.out.println("这是一个时间类型（LocalDateTime）");
            cell.setCellValue((LocalDateTime) fieldValue);
        } else if (fieldType == LocalDate.class) {
            System.out.println("这是一个日期类型（LocalDate）");
            cell.setCellValue((LocalDate) fieldValue);
        } else if (fieldType == java.sql.Date.class) {
            System.out.println("这是一个sql时间类型（sql.Date）");
            cell.setCellValue((java.sql.Date) fieldValue);
        } else if (fieldType == java.sql.Time.class) {
            System.out.println("这是一个sql时间类型（sql.Time）");
            cell.setCellValue((java.sql.Time) fieldValue);
        } else if (fieldType == java.sql.Timestamp.class) {
            System.out.println("这是一个sql时间类型（sql.Timestamp）");
            cell.setCellValue((java.sql.Timestamp) fieldValue);
        } else {
            System.out.println("未知类型：" + fieldType.getName());
        }
        if (!excel.dateFormat().isEmpty()) {
            dataCellStyle.setDataFormat(workbook.createDataFormat().getFormat(excel.dateFormat()));
        }
        cell.setCellStyle(dataCellStyle);
    }


    private <T> void writeDtoData(Sheet sheet, List<T> dataList, CellStyle cellStyle) {
        for (int i = 0; i < dataList.size(); i++) {
            SXSSFRow row = (SXSSFRow) sheet.createRow(i + 1);
            T t = dataList.get(i);
            Class<?> dataClass = t.getClass();
            List<Field> fields = filterNoExcelField(dataClass.getDeclaredFields());
            for (int j = 0; j < fields.size(); j++) {
                Field field = fields.get(j);
                field.setAccessible(true);
                Annotation[] annotations = field.getAnnotations();
                for (Annotation annotation : annotations) {
                    if (annotation.annotationType() == Excel.class) {
                        try {
                            SXSSFCell cell = row.createCell(j);
                            Excel excel = (Excel) annotation;
                            // 设置行的高度
                            row.setHeightInPoints(excel.height());
                            // 获取字段具体值
                            Object fieldValue = field.get(t);
//                            System.out.println(fieldType.getTypeName() + ":" + field.getName() + ":" + fieldValue);
                            if (fieldValue != null) {
                                parseDataType(field, fieldValue, excel, cell, cellStyle);
                            }
//                            cell.setCellStyle(cellStyle);
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
        OutputStream outputStream = Files.newOutputStream(Paths.get(String.format("D://excel/%s.xlsx", IdUtil.simpleUUID())));


//        List<String> headList = Arrays.asList("工号", "姓名", "年龄", "性别", "组织部门");
//        List<List<String>> dataList = Arrays.asList(
//                Arrays.asList("1", "张三", "23", "男", "财务部"),
//                Arrays.asList("2", "李四", "24", "男", "财务部"),
//                Arrays.asList("3", "王五", "25", "男", "财务部")
//        );
//        poiExcelUtil.exportExcel(headList, dataList, outputStream);


//        List<User> userList = new ArrayList<>();
//        userList.add(new User(null, "张三", 0, "a", "123@qq.com", "123456789", "1", null, null));
//        userList.add(new User(null, "李四", 0, "a", "123@qq.com", "123456789", "1", null, null));
//        userList.add(new User(null, "王五", 0, "a", "123@qq.com", "123456789", "1", null, null));
//        poiExcelUtil.exportExcel(userList, outputStream);


        List<Test> testList = new ArrayList<>();

        testList.add(new Test(1, 2, "string", 3.123456, 1.56789f, 123456L, (short) 123, null, true, '1', new Date(), new java.sql.Date(new Date().getTime()), new Time(new Date().getTime()), new Timestamp(new Date().getTime()), LocalDate.now(), LocalTime.now(), LocalDateTime.now()));

        poiExcelUtil.exportExcel(testList, outputStream);
        System.out.println(new Date());
    }


}
