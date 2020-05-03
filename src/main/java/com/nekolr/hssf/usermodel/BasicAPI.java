package com.nekolr.hssf.usermodel;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Calendar;
import java.util.Date;

public class BasicAPI {

    public static void setCellValue(HSSFCell cell) {
        cell.setCellValue(true);
        // double
        cell.setCellValue(233);
        cell.setCellValue(new Date());
        cell.setCellValue("hello world");
        // JDK 8 LocalDate LocalDateTime
        cell.setCellValue(LocalDate.now());
        cell.setCellValue(LocalDateTime.now());
        cell.setCellValue(new HSSFRichTextString("rich text string"));
        cell.setCellValue(Calendar.getInstance());
        cell.setCellErrorValue(FormulaError.NUM);
    }

    public static void setCellValueWithDateTimeFormat(HSSFWorkbook workbook, HSSFCell cell) {
        CreationHelper creationHelper = workbook.getCreationHelper();
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));

        // 使用内建的格式
        cellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("m/d/yy h:mm"));

        cell.setCellValue(new Date());
        cell.setCellStyle(cellStyle);
    }

    public static String getCellValue(HSSFCell cell) {
        String value;
        switch (cell.getCellType()) {
            case FORMULA:
                value = "FORMULA value=" + cell.getCellFormula();
                break;
            case NUMERIC:
                value = "NUMERIC value=" + cell.getNumericCellValue();
                break;
            case STRING:
                value = "STRING value=" + cell.getStringCellValue();
                break;
            case BLANK:
                value = "<BLANK>";
                break;
            case BOOLEAN:
                value = "BOOLEAN value-" + cell.getBooleanCellValue();
                break;
            case ERROR:
                value = "ERROR value=" + cell.getErrorCellValue();
                break;
            default:
                value = "UNKNOWN value of type " + cell.getCellType();
        }
        return value;
    }

    public static void main(String[] args) throws IOException {
        // 两种方式拿到文件流
        InputStream inputStream = new FileInputStream("d:\\workbook.xls");
//        POIFSFileSystem fileSystem = new POIFSFileSystem(new FileInputStream("test.xls"));

        // Excel 工作簿对象
        HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
//        HSSFWorkbook workbook = new HSSFWorkbook(fileSystem);

        // 通过工厂类创建工作簿对象，工厂类会根据文件头部的信息来判断这是一个 xls 文件还是一个 xlsx 文件，
        // 这里的 Workbook 是所有工作簿对象的接口
//        Workbook workbook = WorkbookFactory.create(inputStream);


        // Sheet 工作表对象
        HSSFSheet sheet = workbook.getSheet("sheet1");
//        HSSFSheet sheet1 = workbook.getSheetAt(0);

        // Row 工作表的行
        HSSFRow row = sheet.getRow(0);

        // Cell 工作表指定的单元格
        HSSFCell cell = row.getCell(0);

        // 单元格的样式
        HSSFCellStyle style = cell.getCellStyle();

        // 单元格的类型，是数字、公式、字符串还是
        CellType cellType = cell.getCellType();

        // 获取公式，比如：A1 + B1
//        String formula = cell.getCellFormula();

        // 获取单元格上的批注
//        HSSFComment comment = cell.getCellComment();


        // 写文件
        OutputStream outputStream = new FileOutputStream("d:\\workbook.xls");
        workbook.write(outputStream);

        inputStream.close();
        outputStream.close();

    }
}
