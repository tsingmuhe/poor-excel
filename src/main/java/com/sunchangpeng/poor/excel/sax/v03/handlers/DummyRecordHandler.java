package com.sunchangpeng.poor.excel.sax.v03.handlers;

import com.sunchangpeng.poor.excel.sax.v03.XlsReadContext;
import org.apache.poi.hssf.eventusermodel.dummyrecord.LastCellOfRowDummyRecord;
import org.apache.poi.hssf.record.Record;

import java.util.LinkedHashMap;

public class DummyRecordHandler implements IgnorableXlsRecordHandler {
    @Override
    public void processRecord(XlsReadContext context, Record record) {
        if (record instanceof LastCellOfRowDummyRecord) {
            // End of this row
            LastCellOfRowDummyRecord drd = (LastCellOfRowDummyRecord) record;
            context.setLastRowIndex(drd.getRow());

            processLastCellOfRowDummyRecord(context, drd.getRow());
        }
    }

    private void processLastCellOfRowDummyRecord(XlsReadContext context, int row) {
        if (context.matchRowFilter(context.getCurrentSheetIndex(), row)) {
            if (!(context.getConfig().isSkipEmptyRow() && context.isEmptyRow())) {
                context.handle(context.getCurrentSheetIndex(), row, context.getCellMap());
            }

            context.setCellMap(new LinkedHashMap<>());
        }
    }
}
