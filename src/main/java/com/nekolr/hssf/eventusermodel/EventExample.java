package com.nekolr.hssf.eventusermodel;

import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.record.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

public class EventExample implements HSSFListener {

    private SSTRecord sstRecord;

    @Override
    public void processRecord(org.apache.poi.hssf.record.Record record) {
        switch (record.getSid()) {
            // Workbook 或者 sheet 的开始
            case BOFRecord.sid:
                BOFRecord bof = (BOFRecord) record;
                if (bof.getType() == BOFRecord.TYPE_WORKBOOK) {
                    System.out.println("workbook 的开始");
                } else if (bof.getType() == BOFRecord.TYPE_WORKSHEET) {
                    System.out.println("sheet 的开始");
                }
                break;
            // sheet 名称
            case BoundSheetRecord.sid:
                BoundSheetRecord bsr = (BoundSheetRecord) record;
                System.out.println("sheet name: " + bsr.getSheetname());
                break;
            // 行
            case RowRecord.sid:
                RowRecord rowRecord = (RowRecord) record;
                System.out.println("Row found, first column at "
                        + rowRecord.getFirstCol() + " last column at " + rowRecord.getLastCol());
                break;
            // 数字单元格
            case NumberRecord.sid:
                NumberRecord numberRecord = (NumberRecord) record;
                System.out.println("Cell found with value " + numberRecord.getValue()
                        + " at row " + numberRecord.getRow() + " and column " + numberRecord.getColumn());
                break;
            // SSTRecords store a array of unique strings used in Excel.
            // 存储所有具有唯一性的字符串
            case SSTRecord.sid:
                sstRecord = (SSTRecord) record;
                for (int k = 0; k < sstRecord.getNumUniqueStrings(); k++) {
                    System.out.println("String table value " + k + " = " + sstRecord.getString(k));
                }
                break;
            // 字符串单元格
            case LabelSSTRecord.sid:
                LabelSSTRecord labelSSTRecord = (LabelSSTRecord) record;
                System.out.println("String cell found with value "
                        + sstRecord.getString(labelSSTRecord.getSSTIndex()));
                break;
        }
    }

    public static void main(String[] args) throws IOException {
        try (FileInputStream fin = new FileInputStream("d:\\workbook.xls")) {
            try (POIFSFileSystem poifsFileSystem = new POIFSFileSystem(fin)) {
                // get the Workbook (excel part) stream in a InputStream
                try (InputStream din = poifsFileSystem.createDocumentInputStream("Workbook")) {
                    // construct out HSSFRequest object
                    HSSFRequest req = new HSSFRequest();
                    // lazy listen for ALL records with the listener shown above
                    req.addListenerForAllRecords(new EventExample());
                    // create our event factory
                    HSSFEventFactory factory = new HSSFEventFactory();
                    // process our events based on the document input stream
                    factory.processEvents(req, din);
                }
            }
        }
        System.out.println("done.");
    }
}
