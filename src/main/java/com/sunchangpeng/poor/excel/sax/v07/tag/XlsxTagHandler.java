package com.sunchangpeng.poor.excel.sax.v07.tag;

import com.sunchangpeng.poor.excel.sax.v07.XlsxReadContext;
import org.xml.sax.Attributes;

public interface XlsxTagHandler {
    default void startElement(XlsxReadContext context, String name, Attributes attributes) {
    }

    default void characters(XlsxReadContext context, char[] ch, int start, int length) {
    }

    default void endElement(XlsxReadContext context, String name) {
    }
}
