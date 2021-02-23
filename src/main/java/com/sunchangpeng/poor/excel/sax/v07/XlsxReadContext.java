package com.sunchangpeng.poor.excel.sax.v07;

import com.sunchangpeng.poor.excel.cell.CellData;
import com.sunchangpeng.poor.excel.sax.ExcelSaxReaderConfig;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;

import java.util.Map;

public class XlsxReadContext {
    @Getter
    private final ExcelSaxReaderConfig config;

    //context
    @Getter
    @Setter
    private StylesTable stylesTable;
    @Getter
    @Setter
    private SharedStringsTable sharedStringsTable;
    @Getter
    private final int currentSheetIndex;

    //runtime
    @Getter
    @Setter
    private int lastRowIndex = -1;
    @Getter
    @Setter
    private Map<Integer, CellData> cellMap;
    @Getter
    @Setter
    private int lastColumnIndex = -1;
    @Getter
    @Setter
    private XlsxCellRuntime cellRuntime;

    public XlsxReadContext(ExcelSaxReaderConfig config, int currentSheetIndex) {
        this.config = config;
        this.currentSheetIndex = currentSheetIndex;
    }

    public boolean isEmptyRow() {
        return this.cellMap == null || this.cellMap.isEmpty();
    }

    public boolean matchRowFilter(Integer t, Integer u) {
        return this.config.getRowFilter().test(t, u);
    }

    public void handle(int sheetIndex, long rowIndex, Map<Integer, CellData> cellMap) {
        this.config.getRowHandler().handle(sheetIndex, rowIndex, cellMap);
    }
}
