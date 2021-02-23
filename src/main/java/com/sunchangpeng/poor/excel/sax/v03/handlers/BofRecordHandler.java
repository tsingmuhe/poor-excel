package com.sunchangpeng.poor.excel.sax.v03.handlers;

import com.sunchangpeng.poor.excel.sax.v03.XlsReadContext;
import com.sunchangpeng.poor.excel.sax.v03.XlsRecordHandler;
import org.apache.poi.hssf.record.BOFRecord;
import org.apache.poi.hssf.record.Record;

public class BofRecordHandler implements XlsRecordHandler {
    @Override
    public void processRecord(XlsReadContext context, Record record) {
        // the BOFRecord can represent either the beginning of a sheet or the workbook
        BOFRecord br = (BOFRecord) record;
        if (br.getType() == BOFRecord.TYPE_WORKSHEET) {
            context.setCurrentSheetIndex(context.getCurrentSheetIndex() + 1);
        }
    }
}
