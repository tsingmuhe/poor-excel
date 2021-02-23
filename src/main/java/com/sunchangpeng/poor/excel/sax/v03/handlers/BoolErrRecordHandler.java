package com.sunchangpeng.poor.excel.sax.v03.handlers;

import com.sunchangpeng.poor.excel.cell.BooleanCellData;
import com.sunchangpeng.poor.excel.sax.v03.XlsReadContext;
import org.apache.poi.hssf.record.BoolErrRecord;
import org.apache.poi.hssf.record.Record;

public class BoolErrRecordHandler implements IgnorableXlsRecordHandler {
    @Override
    public void processRecord(XlsReadContext context, Record record) {
        BoolErrRecord brd = (BoolErrRecord) record;
        if (!context.matchRowFilter(context.getCurrentSheetIndex(), brd.getRow())) {
            return;
        }

        context.getCellMap().put((int) brd.getColumn(), new BooleanCellData().setValue(brd.getBooleanValue()));
    }
}
