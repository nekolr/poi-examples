package com.nekolr.hssf.usermodel;

import org.apache.poi.hssf.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;

public class CellComment {
    public static void main(String[] args) throws IOException {
        try (HSSFWorkbook wb = new HSSFWorkbook()) {

            HSSFSheet sheet = wb.createSheet("sheet1");
            HSSFCell cell = sheet.createRow(0).createCell(0);

            HSSFPatriarch drawingPatriarch = sheet.createDrawingPatriarch();
            HSSFComment comment = drawingPatriarch.createCellComment(new HSSFClientAnchor(0, 0, 0, 0, (short) 0, 3, (short) 4, 7));
            comment.setAuthor("saber");
            comment.setString(new HSSFRichTextString("This is a comment"));
            cell.setCellComment(comment);

            try (FileOutputStream fileOut = new FileOutputStream("d:\\workbook.xls")) {
                wb.write(fileOut);
            }
        }
    }
}
