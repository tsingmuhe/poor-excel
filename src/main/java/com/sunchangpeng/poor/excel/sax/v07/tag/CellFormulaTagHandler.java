package com.sunchangpeng.poor.excel.sax.v07.tag;

import com.sunchangpeng.poor.excel.sax.v07.XlsxCellRuntime;
import com.sunchangpeng.poor.excel.sax.v07.XlsxReadContext;
import org.xml.sax.Attributes;

public class CellFormulaTagHandler implements IgnorableXlsxTagHandler {
    @Override
    public void startElement(XlsxReadContext context, String name, Attributes attributes) {
        XlsxCellRuntime cellRuntime = context.getCellRuntime();
        cellRuntime.setFormula(true);
        cellRuntime.setOriginFormulaValue(new StringBuilder());
    }

    @Override
    public void characters(XlsxReadContext context, char[] ch, int start, int length) {
        context.getCellRuntime().getOriginFormulaValue().append(ch, start, length);
    }

    @Override
    public void endElement(XlsxReadContext context, String name) {
        XlsxCellRuntime cellRuntime = context.getCellRuntime();
        cellRuntime.setFormulaValue(cellRuntime.getOriginFormulaValue().toString());
    }
}
