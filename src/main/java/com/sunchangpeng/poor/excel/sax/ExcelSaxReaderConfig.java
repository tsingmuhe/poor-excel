package com.sunchangpeng.poor.excel.sax;

import com.sunchangpeng.poor.excel.sax.v03.Excel03SaxReader;
import com.sunchangpeng.poor.excel.sax.v07.Excel07SaxReader;
import com.sunchangpeng.poor.excel.utils.ExcelFileUtil;
import lombok.Getter;

import java.io.File;

public class ExcelSaxReaderConfig {
    @Getter
    private final File file;

    @Getter
    private boolean use1904windowing = false;
    @Getter
    private boolean skipEmptyRow = true;
    @Getter
    private boolean autoTrim = false;
    @Getter
    private int targetSheet = -1;

    @Getter
    private RowHandler rowHandler;
    @Getter
    private RowFilter rowFilter = (sheet, row) -> true;

    public ExcelSaxReaderConfig(File file) {
        this.file = file;
    }

    public ExcelSaxReaderConfig use1904windowing(boolean use1904windowing) {
        this.use1904windowing = use1904windowing;
        return this;
    }

    public ExcelSaxReaderConfig skipEmptyRow(boolean skipEmptyRow) {
        this.skipEmptyRow = skipEmptyRow;
        return this;
    }

    public ExcelSaxReaderConfig autoTrim(boolean autoTrim) {
        this.autoTrim = autoTrim;
        return this;
    }

    public ExcelSaxReaderConfig sheet(int targetSheet) {
        this.targetSheet = targetSheet;
        return this;
    }

    public ExcelSaxReaderConfig rowFilter(RowFilter rowFilter) {
        this.rowFilter = rowFilter;
        return this;
    }

    public ExcelSaxReaderConfig rowHandler(RowHandler rowHandler) {
        this.rowHandler = rowHandler;
        return this;
    }

    public void read() {
        if (file == null) {
            throw new NullPointerException();
        }

        ExcelSaxReader reader = ExcelFileUtil.isXlsx(file) ? new Excel07SaxReader() : new Excel03SaxReader();
        reader.read(this);
    }
}
