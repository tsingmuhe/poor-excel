package com.sunchangpeng.poor.excel.sax.v03.handlers;

import com.sunchangpeng.poor.excel.sax.v03.XlsReadContext;
import com.sunchangpeng.poor.excel.sax.v03.XlsRecordHandler;
import org.apache.poi.hssf.record.BoundSheetRecord;
import org.apache.poi.hssf.record.Record;

public class BoundSheetRecordHandler implements XlsRecordHandler {
    @Override
    public void processRecord(XlsReadContext context, Record record) {
        context.getBoundSheetRecords().add((BoundSheetRecord) record);
    }
}
