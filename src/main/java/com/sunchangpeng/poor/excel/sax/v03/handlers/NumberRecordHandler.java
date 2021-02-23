package com.sunchangpeng.poor.excel.sax.v03.handlers;

import com.sunchangpeng.poor.excel.cell.DateCellData;
import com.sunchangpeng.poor.excel.cell.NumberCellData;
import com.sunchangpeng.poor.excel.sax.v03.XlsReadContext;
import com.sunchangpeng.poor.excel.utils.ExcelDateUtil;
import org.apache.poi.hssf.record.NumberRecord;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.ss.usermodel.BuiltinFormats;

import java.math.BigDecimal;

import static org.apache.poi.ss.usermodel.DateUtil.getJavaDate;

public class NumberRecordHandler implements IgnorableXlsRecordHandler {
    @Override
    public void processRecord(XlsReadContext context, Record record) {
        NumberRecord nrd = (NumberRecord) record;
        if (!context.matchRowFilter(context.getCurrentSheetIndex(), nrd.getRow())) {
            return;
        }

        int dataFormat = context.getFormatTrackingHSSFListener().getFormatIndex(nrd);
        String dataFormatString = BuiltinFormats.getBuiltinFormat(dataFormat);
        if (dataFormatString == null) {
            dataFormatString = context.getFormatTrackingHSSFListener().getFormatString(nrd);
        }

        if (ExcelDateUtil.isADateFormat(dataFormat, dataFormatString)) {
            context.getCellMap().put((int) nrd.getColumn(), new DateCellData()
                    .setValue(getJavaDate(nrd.getValue(), context.getConfig().isUse1904windowing())));
        } else {
            context.getCellMap().put((int) nrd.getColumn(), new NumberCellData()
                    .setValue(BigDecimal.valueOf(nrd.getValue())));
        }
    }
}
