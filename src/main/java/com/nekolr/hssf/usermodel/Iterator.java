package com.nekolr.hssf.usermodel;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.FileInputStream;
import java.io.IOException;

public class Iterator {
    public static void main(String[] args) throws IOException {

        try (HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream("d:\\workbook.xls"))) {

            HSSFSheet sheet = wb.getSheetAt(0);

            // 迭代器
            for (java.util.Iterator<Row> rowIterator = sheet.rowIterator(); rowIterator.hasNext(); ) {
                Row row = rowIterator.next();
                for (java.util.Iterator<Cell> cellIterator = row.cellIterator(); cellIterator.hasNext(); ) {
                    Cell cell = cellIterator.next();
                    // do something
                }
            }

            // 增强 for 循环
            for (Row row : sheet) {
                for (Cell cell : row) {
                    // do something
                }
            }
        }


    }
}
