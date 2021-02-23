package com.sunchangpeng.poor.excel.sax.v07;

import com.sunchangpeng.poor.excel.sax.v07.tag.IgnorableXlsxTagHandler;
import com.sunchangpeng.poor.excel.sax.v07.tag.TagName;
import com.sunchangpeng.poor.excel.sax.v07.tag.XlsxTagHandler;
import org.xml.sax.Attributes;
import org.xml.sax.helpers.DefaultHandler;

import java.util.Deque;
import java.util.LinkedList;

public class XlsxSaxAnalyser extends DefaultHandler {
    private final Deque<String> tagDeque = new LinkedList<>();
    private final XlsxReadContext context;

    public XlsxSaxAnalyser(XlsxReadContext context) {
        this.context = context;
    }

    @Override
    public void startElement(String uri, String localName, String name, Attributes attributes) {
        XlsxTagHandler handler = TagName.getTagHandler(name);
        if (handler == null) {
            return;
        }

        tagDeque.push(name);

        if (isIgnorable(handler)) {
            return;
        }

        handler.startElement(context, name, attributes);
    }

    @Override
    public void characters(char[] ch, int start, int length) {
        String currentTag = tagDeque.peek();
        if (currentTag == null) {
            return;
        }

        XlsxTagHandler handler = TagName.getTagHandler(currentTag);
        if (handler == null) {
            return;
        }

        if (isIgnorable(handler)) {
            return;
        }

        handler.characters(context, ch, start, length);
    }

    @Override
    public void endElement(String uri, String localName, String name) {
        XlsxTagHandler handler = TagName.getTagHandler(name);
        if (handler == null) {
            return;
        }

        if (!isIgnorable(handler)) {
            handler.endElement(context, name);
        }

        tagDeque.pop();
    }

    private boolean isIgnorable(XlsxTagHandler handler) {
        if (!(handler instanceof IgnorableXlsxTagHandler)) {
            return false;
        }

        return !context.matchRowFilter(context.getCurrentSheetIndex(), context.getLastRowIndex());
    }
}
