package com.sunchangpeng.poor.excel.sax.v07;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class XlsxCellRuntime {
    private XlsxCellDataType cellDataType;

    private int dataFormat;
    private String dataFormatString;

    private boolean formula;
    private StringBuilder originFormulaValue;
    private String formulaValue;

    private StringBuilder originValue;
    private String value;
}
