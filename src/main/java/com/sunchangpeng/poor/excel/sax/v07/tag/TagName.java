package com.sunchangpeng.poor.excel.sax.v07.tag;

public enum TagName {
    /**
     * row
     */
    row(new RowTagHandler()),
    /**
     * cell
     */
    c(new CellTagHandler()),
    /**
     * cell value
     */
    v(new CellValueTagHandler()),
    /**
     * for Formula
     */
    f(new CellFormulaTagHandler()),
    /**
     * for inlineStr
     */
    t(new CellInlineStringValueTagHandler());

    private final XlsxTagHandler xlsxTagHandler;

    TagName(XlsxTagHandler xlsxTagHandler) {
        this.xlsxTagHandler = xlsxTagHandler;
    }

    public static TagName getByTagName(String tagName) {
        if (row.name().equals(tagName)) {
            return row;
        }

        if (c.name().equals(tagName)) {
            return c;
        }

        if (v.name().equals(tagName)) {
            return v;
        }

        if (f.name().equals(tagName)) {
            return f;
        }


        if (t.name().equals(tagName)) {
            return t;
        }

        return null;
    }

    public static XlsxTagHandler getTagHandler(String tagName) {
        TagName target = getByTagName(tagName);
        return target == null ? null : target.getXlsxTagHandler();
    }

    public XlsxTagHandler getXlsxTagHandler() {
        return xlsxTagHandler;
    }
}
