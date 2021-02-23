package com.sunchangpeng.poor.excel.sax.v07.tag;

import com.sunchangpeng.poor.excel.sax.v07.XlsxReadContext;
import org.xml.sax.Attributes;

import java.util.LinkedHashMap;

public class RowTagHandler implements XlsxTagHandler {
    @Override
    public void startElement(XlsxReadContext context, String name, Attributes attributes) {
        String rValue = AttributeName.r.getValue(attributes);
        if (rValue == null) {
            return;
        }

        int currentRowIndex = Integer.parseInt(rValue) - 1;
        int lastRowIndex = context.getLastRowIndex();
        if (!skipEmptyRow(context, lastRowIndex, currentRowIndex)) {
            handleEmptyRow(context, lastRowIndex, currentRowIndex);
        }

        context.setLastRowIndex(currentRowIndex);
        context.setCellMap(new LinkedHashMap<>());
    }

    private boolean skipEmptyRow(XlsxReadContext context, int lastRowIndex, int currentRowIndex) {
        if (context.getConfig().isSkipEmptyRow()) {
            return true;
        }

        return lastRowIndex + 1 >= currentRowIndex;
    }

    private void handleEmptyRow(XlsxReadContext context, int lastRowIndex, int currentRowIndex) {
        for (int i = lastRowIndex + 1; i < currentRowIndex; i++) {
            if (context.matchRowFilter(context.getCurrentSheetIndex(), i)) {
                context.handle(context.getCurrentSheetIndex(), i, new LinkedHashMap<>());
            }
        }
    }

    @Override
    public void endElement(XlsxReadContext context, String name) {
        if (context.matchRowFilter(context.getCurrentSheetIndex(), context.getLastRowIndex())) {
            if (!(context.getConfig().isSkipEmptyRow() && context.isEmptyRow())) {
                context.handle(context.getCurrentSheetIndex(), context.getLastRowIndex(), context.getCellMap());
            }
        }

        context.setCellMap(null);
    }
}
