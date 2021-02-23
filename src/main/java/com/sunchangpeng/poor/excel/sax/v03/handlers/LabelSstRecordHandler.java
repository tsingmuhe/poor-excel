package com.sunchangpeng.poor.excel.sax.v03.handlers;

import com.sunchangpeng.poor.excel.cell.StringCellData;
import com.sunchangpeng.poor.excel.sax.v03.XlsReadContext;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.record.LabelSSTRecord;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.record.SSTRecord;

public class LabelSstRecordHandler implements IgnorableXlsRecordHandler {
    @Override
    public void processRecord(XlsReadContext context, Record record) {
        // shared string table
        LabelSSTRecord lsrd = (LabelSSTRecord) record;
        if (!context.matchRowFilter(context.getCurrentSheetIndex(), lsrd.getRow())) {
            return;
        }

        SSTRecord sstRecord = context.getSstRecord();
        if (sstRecord == null) {
            return;
        }

        String data = sstRecord.getString(lsrd.getSSTIndex()).toString();
        if (data == null) {
            return;
        }

        context.getCellMap().put((int) lsrd.getColumn(), new StringCellData()
                .setValue(context.getConfig().isAutoTrim() ? StringUtils.trim(data) : data));
    }
}
