package com.nekolr.hssf.usermodel;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileOutputStream;
import java.io.IOException;

public class MergedCells {
    public static void main(String[] args) throws IOException {
        try (HSSFWorkbook wb = new HSSFWorkbook()) {
            HSSFSheet sheet = wb.createSheet("sheet1");

            HSSFRow row = sheet.createRow(0);
            HSSFCell cell = row.createCell(0);
            cell.setCellValue("This is a test of merging");

            // 合并单元格，第一行的第一个单元格与第二个单元格合并
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 1));

            try (FileOutputStream fileOut = new FileOutputStream("d:\\workbook.xls")) {
                wb.write(fileOut);
            }
        }
    }
}
