package com.sunchangpeng.poor.excel.sax.v03.handlers;

import com.sunchangpeng.poor.excel.sax.v03.XlsReadContext;
import com.sunchangpeng.poor.excel.sax.v03.XlsRecordHandler;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.record.SSTRecord;

public class SstRecordHandler implements XlsRecordHandler {
    @Override
    public void processRecord(XlsReadContext context, Record record) {
        // shared string table
        context.setSstRecord((SSTRecord) record);
    }
}
