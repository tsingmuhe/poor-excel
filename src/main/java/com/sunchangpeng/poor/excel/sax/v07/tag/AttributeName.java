package com.sunchangpeng.poor.excel.sax.v07.tag;

import org.xml.sax.Attributes;

public enum AttributeName {
    /**
     * row num or col num
     */
    r,
    /**
     * ST（StylesTable）index
     */
    s,
    /**
     * cell type
     */
    t;

    public String getValue(Attributes attributes) {
        return attributes.getValue(name());
    }
}
