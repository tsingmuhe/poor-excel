package com.sunchangpeng.poor.excel.sax.v03.handlers;

import com.sunchangpeng.poor.excel.cell.StringCellData;
import com.sunchangpeng.poor.excel.sax.v03.XlsReadContext;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.record.LabelRecord;
import org.apache.poi.hssf.record.Record;

public class LabelRecordHandler implements IgnorableXlsRecordHandler {
    @Override
    public void processRecord(XlsReadContext context, Record record) {
        //direct text
        LabelRecord lrd = (LabelRecord) record;
        if (!context.matchRowFilter(context.getCurrentSheetIndex(), lrd.getRow())) {
            return;
        }

        context.getCellMap().put((int) lrd.getColumn(), new StringCellData()
                .setValue(context.getConfig().isAutoTrim() ? StringUtils.trim(lrd.getValue()) : lrd.getValue()));
    }
}
