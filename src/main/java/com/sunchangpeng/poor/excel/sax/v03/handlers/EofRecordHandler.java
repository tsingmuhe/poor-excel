package com.sunchangpeng.poor.excel.sax.v03.handlers;

import com.sunchangpeng.poor.excel.sax.v03.XlsReadContext;
import com.sunchangpeng.poor.excel.sax.v03.XlsRecordHandler;
import org.apache.poi.hssf.record.Record;

import java.util.LinkedHashMap;

public class EofRecordHandler implements XlsRecordHandler {
    @Override
    public void processRecord(XlsReadContext context, Record record) {
        // the EOFRecord can represent the end of a sheet

        // Sometimes tables lack the end record of the last column
        if (!context.isEmptyRow()) {
            if (context.matchRowFilter(context.getCurrentSheetIndex(), context.getLastRowIndex() + 1)) {
                context.handle(context.getCurrentSheetIndex(), context.getLastRowIndex() + 1, context.getCellMap());
            }
            context.setCellMap(new LinkedHashMap<>());
        }

        context.getConfig().getRowHandler().endSheet(context.getCurrentSheetIndex());
    }
}
