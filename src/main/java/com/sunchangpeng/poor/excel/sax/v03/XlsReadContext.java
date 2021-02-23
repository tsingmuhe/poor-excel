package com.sunchangpeng.poor.excel.sax.v03;

import com.sunchangpeng.poor.excel.cell.CellData;
import com.sunchangpeng.poor.excel.sax.ExcelSaxReaderConfig;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.hssf.eventusermodel.FormatTrackingHSSFListener;
import org.apache.poi.hssf.record.BoundSheetRecord;
import org.apache.poi.hssf.record.SSTRecord;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

public class XlsReadContext {
    //config
    @Getter
    private final ExcelSaxReaderConfig config;

    //context
    @Getter
    @Setter
    private FormatTrackingHSSFListener formatTrackingHSSFListener;
    @Getter
    @Setter
    private HSSFWorkbook hssfWorkbook;
    @Getter
    @Setter
    private SSTRecord sstRecord;
    @Getter
    private final List<BoundSheetRecord> boundSheetRecords = new ArrayList<>();

    //runtime
    @Getter
    @Setter
    private int currentSheetIndex = -1;
    @Getter
    @Setter
    private int lastRowIndex = -1;
    @Getter
    @Setter
    private Map<Integer, CellData> cellMap = new LinkedHashMap<>();

    @Getter
    @Setter
    private Integer tempFormulaColumn;
    @Getter
    @Setter
    private String tempFormulaValue;

    public XlsReadContext(ExcelSaxReaderConfig config) {
        this.config = config;
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
