package com.nekolr.hssf.usermodel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.WorkbookUtil;

import java.io.FileOutputStream;
import java.io.IOException;

public class NewSheet {
    public static void main(String[] args) throws IOException {
        try (HSSFWorkbook wb = new HSSFWorkbook()) {
            wb.createSheet("sheet1");
            // 使用默认名称创建
            wb.createSheet();
            final String name = "sheet2";
            // 修改 sheet 名称
            wb.setSheetName(1, WorkbookUtil.createSafeSheetName(name));
            try (FileOutputStream fileOut = new FileOutputStream("d:\\workbook.xls")) {
                wb.write(fileOut);
            }
        }
    }
}
