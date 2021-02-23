package com.sunchangpeng.poor.excel.sax;

import com.sunchangpeng.poor.excel.cell.CellData;

import java.util.Map;

@FunctionalInterface
public interface RowHandler {
    void handle(int sheetIndex, long rowIndex, Map<Integer, CellData> cellMap);

    default void endSheet(int sheetIndex) {
        //pass
    }
}
