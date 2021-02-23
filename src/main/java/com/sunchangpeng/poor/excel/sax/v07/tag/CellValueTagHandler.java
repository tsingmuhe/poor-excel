package com.sunchangpeng.poor.excel.sax.v07.tag;

import com.sunchangpeng.poor.excel.sax.v07.XlsxCellDataType;
import com.sunchangpeng.poor.excel.sax.v07.XlsxReadContext;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.RichTextString;

public class CellValueTagHandler implements IgnorableXlsxTagHandler {
    @Override
    public void characters(XlsxReadContext context, char[] ch, int start, int length) {
        context.getCellRuntime().getOriginValue().append(ch, start, length);
    }

    @Override
    public void endElement(XlsxReadContext context, String name) {
        XlsxCellDataType cellDataType = context.getCellRuntime().getCellDataType();
        boolean autoTrim = context.getConfig().isAutoTrim();
        if (XlsxCellDataType.STRING == cellDataType) {
            int idx = Integer.parseInt(context.getCellRuntime().getOriginValue().toString());
            RichTextString richTextString = context.getSharedStringsTable().getItemAt(idx);
            context.getCellRuntime().setValue(autoTrim ? StringUtils.trim(richTextString.getString()) : richTextString.getString());
            return;
        } else if (XlsxCellDataType.EMPTY == cellDataType) {
            context.getCellRuntime().setCellDataType(XlsxCellDataType.NUMBER);
        }

        context.getCellRuntime().setValue(autoTrim ? StringUtils.trim(context.getCellRuntime().getOriginValue().toString()) :
                context.getCellRuntime().getOriginValue().toString());
    }
}
