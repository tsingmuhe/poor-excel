package com.sunchangpeng.poor.excel.sax.v07;

import java.util.HashMap;
import java.util.Map;

public enum XlsxCellDataType {
    /**
     * string
     */
    STRING,
    /**
     * rich String,does not need to be read in the 'sharedStrings.xml'
     */
    RICH_TEXT_STRING,
    /**
     * string,This type of data does not need to be read in the 'sharedStrings.xml'
     */
    DIRECT_STRING,
    /**
     * number
     */
    NUMBER,
    /**
     * boolean
     */
    BOOLEAN,
    /**
     * error
     */
    ERROR,
    /**
     * empty,it's means Empty or Number
     */
    EMPTY;

    private static final Map<String, XlsxCellDataType> TYPE_ROUTING_MAP = new HashMap<>(16);

    static {
        TYPE_ROUTING_MAP.put("s", STRING);
        TYPE_ROUTING_MAP.put("inlineStr", RICH_TEXT_STRING);
        TYPE_ROUTING_MAP.put("str", DIRECT_STRING);
        TYPE_ROUTING_MAP.put("n", NUMBER);
        TYPE_ROUTING_MAP.put("b", BOOLEAN);
        TYPE_ROUTING_MAP.put("e", ERROR);
    }

    public static XlsxCellDataType parseCellDataType(String cellType) {
        if (null == cellType) {
            return EMPTY;
        }
        return TYPE_ROUTING_MAP.get(cellType);
    }
}
