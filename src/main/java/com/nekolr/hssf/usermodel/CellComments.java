/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

package com.nekolr.hssf.usermodel;

import org.apache.poi.hssf.usermodel.HSSFComment;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * 单元格的批注，这里是官方提供的例子，批注的大小和坐标设置暂时不清楚是如何计算的
 */
public class CellComments {

    public static void main(String[] args) throws IOException  {
        // 使用 SS API 来操作
        createWorkbook(false, ".xls");
        createWorkbook(true, ".xlsx");
    }

    private static void createWorkbook(boolean xssf, String extension) throws IOException {
        try (Workbook wb = WorkbookFactory.create(xssf)) {

            Sheet sheet = wb.createSheet("sheet1");

            // 使用 CreationHelper 可以不必考虑具体使用 HSSF 还是 XSSF 的 API
            CreationHelper creationHelper = wb.getCreationHelper();

            // 插入图片，绘制图形，设置批注等都要先创建它
            Drawing<?> drawingPatriarch = sheet.createDrawingPatriarch();

            // create a cell in row 3
            Cell cell1 = sheet.createRow(3).createCell(1);
            cell1.setCellValue(creationHelper.createRichTextString("Hello, World"));

            // 控制批注的大小和坐标，这个东西很恶心，暂时不清楚怎么计算的
            ClientAnchor clientAnchor = creationHelper.createClientAnchor();
            clientAnchor.setCol1(4);
            clientAnchor.setRow1(2);
            clientAnchor.setCol2(6);
            clientAnchor.setRow2(5);
            Comment comment1 = drawingPatriarch.createCellComment(clientAnchor);

            // 给批注设置值和作者
            comment1.setString(creationHelper.createRichTextString("We can set comments in POI"));
            comment1.setAuthor("Apache Software Foundation");

            // 给单元格设置批注
            cell1.setCellComment(comment1);

            // create another cell in row 6
            Cell cell2 = sheet.createRow(6).createCell(1);
            cell2.setCellValue(36.6);

            clientAnchor = creationHelper.createClientAnchor();
            clientAnchor.setCol1(4);
            clientAnchor.setRow1(8);
            clientAnchor.setCol2(6);
            clientAnchor.setRow2(11);

            Comment comment2 = drawingPatriarch.createCellComment(clientAnchor);

            if (wb instanceof HSSFWorkbook) {
                // 批注填充背景颜色
                ((HSSFComment) comment2).setFillColor(204, 236, 255);
            }

            RichTextString string = creationHelper.createRichTextString("Normal body temperature");

            // 给批注中的文字设置样式
            Font font = wb.createFont();
            font.setFontName("Arial");
            font.setFontHeightInPoints((short) 10);
            font.setBold(true);
            font.setColor(IndexedColors.RED.getIndex());
            string.applyFont(font);

            comment2.setString(string);
            // 默认情况下批注是隐藏的，此处设置为始终显示
            comment2.setVisible(true);
            comment2.setAuthor("Bill Gates");

            /*
             * 修改批注对应的单元格
             */
            comment2.setRow(6);
            comment2.setColumn(1);

            try (FileOutputStream out = new FileOutputStream("d:\\poi_comment" + extension)) {
                wb.write(out);
            }
        }
    }
}
