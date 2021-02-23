package com.sunchangpeng.poor.excel.sax.v07.tag;

import com.sunchangpeng.poor.excel.sax.v07.XlsxReadContext;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

public class CellInlineStringValueTagHandler implements IgnorableXlsxTagHandler {
    @Override
    public void characters(XlsxReadContext context, char[] ch, int start, int length) {
        context.getCellRuntime().getOriginValue().append(ch, start, length);
    }

    @Override
    public void endElement(XlsxReadContext context, String name) {
        XSSFRichTextString richTextString = new XSSFRichTextString(context.getCellRuntime().getOriginValue().toString());
        context.getCellRuntime().setValue(context.getConfig().isAutoTrim() ? StringUtils.trim(richTextString.getString()) : richTextString.getString());
    }
}
