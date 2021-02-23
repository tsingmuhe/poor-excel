package com.sunchangpeng.poor.excel.sax.v03;

import org.apache.poi.hssf.record.Record;

public interface XlsRecordHandler {
    void processRecord(XlsReadContext context, Record record);
}
