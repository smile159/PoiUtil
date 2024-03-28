package cn.smile.util;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

public class Test {


    public static void test1() throws IOException {
        // 1.创建xlsx的工作簿
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();

        // 2.创建工作表
        XSSFSheet sheet1 = xssfWorkbook.createSheet("我是工作表");

        // 定位到第一行
        XSSFRow row = sheet1.createRow(0);


        // 定位到第一列
        XSSFCell cell = row.createCell(0);

        cell.setCellValue("我是测试的内容");

        // 也可以链式调用
//        sheet1.createRow(0).createCell(0).setCellValue("我是第二行第一列的内容");

        OutputStream outputStream = new FileOutputStream("D://test.xlsx");
        xssfWorkbook.write(outputStream);

        outputStream.close();
        xssfWorkbook.close();

    }

    public static void test2() throws IOException {
        //  1.创建xlsx的工作簿
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook("D://test.xlsx");

        // 2.获取工作表
        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);

        System.out.println(xssfWorkbook.getActiveSheetIndex());

        int numberOfSheets = xssfWorkbook.getNumberOfSheets();
        System.out.println(numberOfSheets);


        XSSFRow row = sheet.getRow(0);

        XSSFCell cell = row.getCell(0);

        String stringCellValue = cell.getStringCellValue();

        System.out.println(stringCellValue);

        xssfWorkbook.close();
    }

    public static void t3() throws IOException {

        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();

        XSSFSheet sheet = xssfWorkbook.createSheet();

        XSSFCreationHelper creationHelper = xssfWorkbook.getCreationHelper();

        short format = creationHelper.createDataFormat().getFormat("yyyy-MM-dd");

        XSSFCellStyle cellStyle = xssfWorkbook.createCellStyle();

        cellStyle.setDataFormat(format);


        XSSFRow row = sheet.createRow(0);

        XSSFCell cell = row.createCell(0);

        // 日期类型不能直接写入，会变成小数
        cell.setCellValue(new Date());
        cell.setCellStyle(cellStyle);


        sheet.createRow(1).createCell(1).setCellValue("我是文本数据");

        XSSFCellStyle cellStyle1 = xssfWorkbook.createCellStyle();

        cellStyle1.setDataFormat(xssfWorkbook.createDataFormat().getFormat("0.00"));

        sheet.createRow(2).createCell(2).setCellValue(123.456);


        sheet.createRow(3).createCell(3).setCellValue(true);


        XSSFCell cell1 = sheet.createRow(4).createCell(4);

        cell1.setCellErrorValue(FormulaError.REF.getCode());


        sheet.createRow(5).createCell(5).setCellValue("123.6666");


        OutputStream outputStream = Files.newOutputStream(Paths.get("D://1.xlsx"));

        xssfWorkbook.write(outputStream);

        outputStream.close();

        xssfWorkbook.close();

    }

    public static void t4() throws IOException {
        //创建数据源
        Object[] title = {"ID", "姓名", "零花钱", "生日", "是否被删除", "空值引用"};
        Object[] s1 = {"tx01", "张三", 66.663D, new Date(), true, FormulaError.NULL};
        Object[] s2 = {"tx02", "李四", 76.882D, new Date(), false, FormulaError.NULL};
        List<Object[]> list = new ArrayList<>();
        list.add(title);
        list.add(s1);
        list.add(s2);

        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();

        XSSFSheet sheet = xssfWorkbook.createSheet();


        for (int i = 0; i < list.size(); i++) {
            // 创建行
            XSSFRow row = sheet.createRow(i);
            for (int j = 0; j < list.get(i).length; j++) {
                XSSFCell cell = row.createCell(j);
                Object o = list.get(i)[j];
                if (o instanceof Date) {
                    XSSFCellStyle cellStyle = xssfWorkbook.createCellStyle();
                    cellStyle.setDataFormat(xssfWorkbook.createDataFormat().getFormat("yyyy-MM-dd"));
                    cell.setCellStyle(cellStyle);
                    cell.setCellValue((Date) o);
                } else if (o instanceof Double) {
                    XSSFCellStyle cellStyle = xssfWorkbook.createCellStyle();
                    cellStyle.setDataFormat(xssfWorkbook.createDataFormat().getFormat("0.00"));
                    cell.setCellStyle(cellStyle);
                    cell.setCellValue((Double) o);
                } else if (o instanceof FormulaError) {
                    cell.setCellErrorValue(((FormulaError) o).getCode());
                } else {
                    cell.setCellValue(o.toString());
                }
            }
        }

        OutputStream outputStream = Files.newOutputStream(Paths.get("D://1.xlsx"));
        xssfWorkbook.write(outputStream);
        outputStream.close();
        xssfWorkbook.close();
        System.out.println("写入完成");
    }

    /*
     * 对齐方式，水平方式和垂直方式
     *
     * */
    public static void alignmentMethod() throws IOException {
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
        XSSFSheet sheet = xssfWorkbook.createSheet();
        sheet.setColumnWidth(1, 20 * 256);
        sheet.setColumnWidth(2, 20 * 256);
        sheet.setColumnWidth(3, 20 * 256);

        XSSFRow row = sheet.createRow(0);
        row.setHeightInPoints(60);

        XSSFCell cell = row.createCell(0);

        cell.setCellValue("测试1");

        XSSFCell cell1 = row.createCell(1);

        cell1.setCellValue("测试2");
        XSSFCellStyle cellStyle = xssfWorkbook.createCellStyle();

        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cell1.setCellStyle(cellStyle);

        XSSFCell cell2 = row.createCell(2);
        cell2.setCellValue("测试3");

        XSSFCellStyle cellStyle2 = xssfWorkbook.createCellStyle();

        cellStyle2.setAlignment(HorizontalAlignment.CENTER);

        cellStyle2.setVerticalAlignment(VerticalAlignment.CENTER);
        cell2.setCellStyle(cellStyle2);
        OutputStream outputStream = Files.newOutputStream(Paths.get("D://alignment.xlsx"));
        xssfWorkbook.write(outputStream);
        outputStream.close();
        xssfWorkbook.close();
        System.out.println("写入完成");
    }

    /*
     * 设置边框颜色
     * */
    public static void borderColor() throws IOException {
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
        XSSFSheet sheet = xssfWorkbook.createSheet();
        XSSFRow row = sheet.createRow(1);
        row.setHeightInPoints(60);
        XSSFCell cell = row.createCell(2);
        cell.setCellValue("哟西");
        XSSFCellStyle cellStyle = xssfWorkbook.createCellStyle();
        // 顶部边框样式
        cellStyle.setBorderTop(BorderStyle.DASH_DOT);
        cellStyle.setBorderLeft(BorderStyle.DASHED);
        cellStyle.setBorderBottom(BorderStyle.DOUBLE);
        cellStyle.setBorderRight(BorderStyle.DASH_DOT_DOT);

        cellStyle.setTopBorderColor(IndexedColors.GREEN.getIndex());
        cellStyle.setLeftBorderColor(IndexedColors.PALE_BLUE.getIndex());
        cellStyle.setBottomBorderColor(IndexedColors.PINK.getIndex());
        cellStyle.setRightBorderColor(IndexedColors.BROWN.getIndex());

        cell.setCellStyle(cellStyle);

        OutputStream outputStream = Files.newOutputStream(Paths.get("D://1.xlsx"));
        xssfWorkbook.write(outputStream);
        outputStream.close();
        xssfWorkbook.close();

    }

    /*
     * 合并单元格
     *
     * */

    public static void mergeCells() throws IOException {
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
        XSSFSheet sheet = xssfWorkbook.createSheet();
        XSSFRow row = sheet.createRow(1);
        XSSFCell cell = row.createCell(1);
        cell.setCellValue("测试合并");

        sheet.addMergedRegion(new CellRangeAddress(1, 1, 1, 5));

        sheet.addMergedRegion(new CellRangeAddress(2, 3, 1, 1));

        OutputStream outputStream = Files.newOutputStream(Paths.get("D://mergeCells.xlsx"));
        xssfWorkbook.write(outputStream);
        outputStream.close();
        xssfWorkbook.close();

    }

    public static void fontStyle() throws IOException {
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
        XSSFSheet sheet = xssfWorkbook.createSheet();
        XSSFRow row = sheet.createRow(1);
        XSSFCell cell = row.createCell(1);

        XSSFFont font = xssfWorkbook.createFont();
        // 设置字体颜色
        font.setColor(IndexedColors.PINK.getIndex());
        // 设置字体粗细 true 粗 false细
        font.setBold(false);
        // 设置字体大小
        font.setFontHeightInPoints((short) 60);
        // 设置字体样式
        font.setFontName("楷体");
        // 设置倾斜
        font.setItalic(true);
        // 设置删除先
        font.setStrikeout(true);
        // 创建样式类
        XSSFCellStyle cellStyle = xssfWorkbook.createCellStyle();
        // 样式设置font
        cellStyle.setFont(font);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("测试字体样式");
        // 输出
        OutputStream outputStream = Files.newOutputStream(Paths.get("D://fontStyle.xlsx"));
        xssfWorkbook.write(outputStream);
        outputStream.close();
        xssfWorkbook.close();
        System.out.println("写入完成");

    }


    /*
     * 获取表头的样式
     * */
    public CellStyle getHeaderCellStyle(SXSSFWorkbook sxssfWorkbook) {
        return null;
    }


    public CellStyle getPublicCellStyle(SXSSFWorkbook sxssfWorkbook) {
        CellStyle cellStyle = sxssfWorkbook.createCellStyle();
        // 设置水平对其方式，设置重置对其方式
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        return cellStyle;
    }


    /*
     *
     * 文件导出，根据位置，创建表头，创建数据
     * headList: 表头集合
     * dataList: 数据集合，其中的值都是String类型的
     * outputStream: 输出流
     * */
    public void exportExcel(List<String> headList, List<List<String>> dataList, OutputStream outputStream) throws IOException {
        if (headList.isEmpty() || dataList.isEmpty()) {
            throw new RuntimeException("参数错误");
        }
        // 获取工作簿
        SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook();
        // 创建工作表
        SXSSFSheet sheet = sxssfWorkbook.createSheet();
        // 创建表头
        SXSSFRow row = sheet.createRow(0);
        for (int i = 0; i < headList.size(); i++) {
            // 创建列
            SXSSFCell cell = row.createCell(i);
            cell.setCellValue(headList.get(i));
            CellStyle publicCellStyle = getPublicCellStyle(sxssfWorkbook);
            cell.setCellStyle(publicCellStyle);
        }
        // 创建数据
        for (int i = 0; i < dataList.size(); i++) {
            // 获取列数据
            List<String> columns = dataList.get(i);
            SXSSFRow dataRow = sheet.createRow(i + 1);
            for (int j = 0; j < columns.size(); j++) {
                SXSSFCell cell = dataRow.createCell(j);
                cell.setCellValue(columns.get(j));
                CellStyle publicCellStyle = getPublicCellStyle(sxssfWorkbook);
                cell.setCellStyle(publicCellStyle);
            }
        }
        sxssfWorkbook.write(outputStream);
        sxssfWorkbook.close();
        System.out.println("文件写入完成");
    }
}
