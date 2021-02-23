package com.sunchangpeng.poor.excel.sax.v03.handlers;

import com.sunchangpeng.poor.excel.cell.StringCellData;
import com.sunchangpeng.poor.excel.sax.v03.XlsReadContext;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.record.StringRecord;

public class StringRecordHandler implements IgnorableXlsRecordHandler {
    @Override
    public void processRecord(XlsReadContext context, Record record) {
        StringRecord srd = (StringRecord) record;
        Integer columnIndex = context.getTempFormulaColumn();
        if (columnIndex == null) {
            return;
        }

        context.getCellMap().put(columnIndex, new StringCellData()
                .setFormula(true)
                .setFormulaValue(context.getTempFormulaValue())
                .setValue(context.getConfig().isAutoTrim() ? StringUtils.trim(srd.getString()) : srd.getString()));

        context.setTempFormulaValue(null);
        context.setTempFormulaColumn(null);
    }
}
