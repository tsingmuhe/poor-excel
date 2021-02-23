package com.sunchangpeng.poor.excel.sax.v03.handlers;

import com.sunchangpeng.poor.excel.cell.BooleanCellData;
import com.sunchangpeng.poor.excel.cell.DateCellData;
import com.sunchangpeng.poor.excel.cell.ErrorCellData;
import com.sunchangpeng.poor.excel.cell.NumberCellData;
import com.sunchangpeng.poor.excel.sax.v03.XlsReadContext;
import com.sunchangpeng.poor.excel.utils.ExcelDateUtil;
import org.apache.poi.hssf.model.HSSFFormulaParser;
import org.apache.poi.hssf.record.FormulaRecord;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaError;

import java.math.BigDecimal;

import static org.apache.poi.ss.usermodel.DateUtil.getJavaDate;

public class FormulaRecordHandler implements IgnorableXlsRecordHandler {
    private static final String ERROR = "#VALUE!";

    @Override
    public void processRecord(XlsReadContext context, Record record) {
        FormulaRecord frd = (FormulaRecord) record;
        if (!context.matchRowFilter(context.getCurrentSheetIndex(), frd.getRow())) {
            return;
        }

        String formulaValue = null;
        try {
            formulaValue = HSSFFormulaParser.toFormulaString(context.getHssfWorkbook(), frd.getParsedExpression());
        } catch (Exception e) {
            // pass
        }

        CellType resultType = CellType.forInt(frd.getCachedResultType());
        switch (resultType) {
            case STRING:
                // Formula result is a string
                // This is stored in the next record
                context.setTempFormulaValue(formulaValue);
                context.setTempFormulaColumn((int) frd.getColumn());
                break;
            case NUMERIC:
                int dataFormat = context.getFormatTrackingHSSFListener().getFormatIndex(frd);
                String dataFormatString = BuiltinFormats.getBuiltinFormat(dataFormat);
                if (dataFormatString == null) {
                    dataFormatString = context.getFormatTrackingHSSFListener().getFormatString(frd);
                }

                if (ExcelDateUtil.isADateFormat(dataFormat, dataFormatString)) {
                    context.getCellMap().put((int) frd.getColumn(), new DateCellData()
                            .setFormula(true).setFormulaValue(formulaValue)
                            .setValue(getJavaDate(frd.getValue(), context.getConfig().isUse1904windowing())));
                } else {
                    context.getCellMap().put((int) frd.getColumn(), new NumberCellData()
                            .setFormula(true).setFormulaValue(formulaValue)
                            .setValue(BigDecimal.valueOf(frd.getValue())));
                }
                break;
            case ERROR:
                context.getCellMap().put((int) frd.getColumn(), new ErrorCellData()
                        .setFormula(true).setFormulaValue(formulaValue)
                        .setValue(FormulaError.forInt(frd.getCachedErrorValue())));
                break;
            case BOOLEAN:
                context.getCellMap().put((int) frd.getColumn(), new BooleanCellData()
                        .setFormula(true).setFormulaValue(formulaValue)
                        .setValue(frd.getCachedBooleanValue()));
                break;
            default:
        }
    }
}
