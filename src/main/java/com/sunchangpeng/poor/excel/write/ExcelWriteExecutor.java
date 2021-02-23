package com.sunchangpeng.poor.excel.write;

import com.sunchangpeng.poor.excel.cell.*;
import com.sunchangpeng.poor.excel.write.style.CellDataStyle;
import com.sunchangpeng.poor.excel.write.style.ExcelWriteStyleSet;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelWriteExecutor implements Closeable {
    /***
     * 03:{@link HSSFWorkbook}
     * 07:{@link SXSSFWorkbook}
     */
    private Workbook workbook;
    private ExcelWriteStyleSet styleSet;
    private Map<Integer, WriteSheetHolder> sheetHolderMap;
    private WriteSheetHolder currentWriteSheetHolder;

    public ExcelWriteExecutor(Workbook workbook) {
        this.workbook = workbook;
        this.currentWriteSheetHolder = getOrCreateSheet(0);
        this.styleSet = new ExcelWriteStyleSet(this.workbook);
    }

    //----global style
    public ExcelWriteExecutor globalStyle(CellDataStyle cellStyle) {
        this.styleSet.resetGlobalCellStyle(cellStyle);
        return this;
    }

    public ExcelWriteExecutor rowStyle(CellDataStyle cellStyle) {
        Row row = getOrCreateRow(getCurrentRow());
        row.setRowStyle(this.styleSet.getOrCreateCellStyle(cellStyle));
        return this;
    }

    public Font createFont() {
        return this.styleSet.createFont();
    }

    //----sheet
    public ExcelWriteExecutor sheet(int sheetIndex) {
        return sheet(sheetIndex, null);
    }

    public ExcelWriteExecutor sheet(int sheetIndex, String sheetName) {
        this.currentWriteSheetHolder = this.getOrCreateSheet(sheetIndex);

        if (StringUtils.isNotBlank(sheetName)) {
            this.renameSheet(sheetName);
        }
        return this;
    }

    //----row
    public int getCurrentRow() {
        return this.currentWriteSheetHolder.getCurrentRowNum().get();
    }

    public ExcelWriteExecutor setCurrentRow(int rowIndex) {
        this.currentWriteSheetHolder.getCurrentRowNum().set(rowIndex);
        return this;
    }

    public ExcelWriteExecutor setCurrentRowToStart() {
        setCurrentRow(0);
        return this;
    }

    public ExcelWriteExecutor setCurrentRowToEnd() {
        setCurrentRow(this.currentWriteSheetHolder.getSheet().getLastRowNum());
        return this;
    }

    public ExcelWriteExecutor passRows(int rows) {
        this.currentWriteSheetHolder.getCurrentRowNum().addAndGet(rows);
        return this;
    }

    //----merge
    public ExcelWriteExecutor mergingCells(int firstColumn, int lastColumn) {
        int currentRowNum = this.currentWriteSheetHolder.getCurrentRowNum().get();
        return mergingCells(currentRowNum, currentRowNum, firstColumn, lastColumn);
    }

    public ExcelWriteExecutor mergingCells(int firstRow, int lastRow, int firstColumn, int lastColumn) {
        final CellRangeAddress cellRangeAddress = new CellRangeAddress(
                firstRow, // first row (0-based)
                lastRow, // last row (0-based)
                firstColumn, // first column (0-based)
                lastColumn // last column (0-based)
        );
        this.currentWriteSheetHolder.getSheet().addMergedRegion(cellRangeAddress);
        return this;
    }

    //----write
    public ExcelWriteExecutor appendRows(List<List<CellData>> rows) {
        for (List<CellData> row : rows) {
            appendRow(row);
        }
        return this;
    }

    public ExcelWriteExecutor appendRow(List<CellData> cellDatas) {
        int rowNum = this.currentWriteSheetHolder.getCurrentRowNum().getAndIncrement();
        Row row = this.currentWriteSheetHolder.getSheet().createRow(rowNum);

        int i = 0;
        for (CellData cellData : cellDatas) {
            writeCell(row, i, cellData);
            i++;
        }
        return this;
    }

    public ExcelWriteExecutor writeCell(int rowId, int colNum, CellData cellData) {
        writeCell(getOrCreateRow(rowId), colNum, cellData);
        return this;
    }

    //----flush
    public ExcelWriteExecutor flush(File file) throws IOException {
        try (FileOutputStream out = new FileOutputStream(file)) {
            flush(out);
        }
        return this;
    }

    /**
     * The user is responsible for closing the outputStream.
     */
    public void flush(OutputStream out) throws IOException {
        this.workbook.write(out);
        out.flush();
    }

    @Override
    public void close() {
        if (this.workbook instanceof SXSSFWorkbook) {
            ((SXSSFWorkbook) this.workbook).dispose();
        }

        IOUtils.closeQuietly(this.workbook, null);

        this.workbook = null;
        this.styleSet = null;
        this.sheetHolderMap = null;
        this.currentWriteSheetHolder = null;
    }

    //-----private
    private WriteSheetHolder getOrCreateSheet(int sheetIndex) {
        if (this.sheetHolderMap == null) {
            this.sheetHolderMap = new HashMap<>();
        }

        WriteSheetHolder writeSheetHolder = this.sheetHolderMap.get(sheetIndex);
        if (writeSheetHolder != null) {
            return writeSheetHolder;
        }

        writeSheetHolder = new WriteSheetHolder(this.workbook, sheetIndex);
        this.sheetHolderMap.put(sheetIndex, writeSheetHolder);
        return writeSheetHolder;
    }

    private void renameSheet(String sheetName) {
        int innerSheetIndex = this.workbook.getSheetIndex(this.currentWriteSheetHolder.getSheet());
        this.workbook.setSheetName(innerSheetIndex, sheetName);
    }

    private Row getOrCreateRow(int rowId) {
        Row row = this.currentWriteSheetHolder.getSheet().getRow(rowId);
        if (row == null) {
            row = this.currentWriteSheetHolder.getSheet().createRow(rowId);
        }
        return row;
    }

    private void writeCell(Row row, int colNum, CellData cellData) {
        if (null == row) {
            return;
        }

        CellStyle cellStyle = (cellData == null || cellData.getStyle() == null) ?
                this.styleSet.getGlobalCellStyle()
                : this.styleSet.getOrCreateCellStyle(cellData.getStyle());

        if (null == cellData || null == cellData.getValue()) {
            if (cellStyle != null) {
                Cell cell = row.createCell(colNum);
                setCellStyle(cell, cellStyle);
            }
            return;
        }

        Cell cell = null;
        if (cellData instanceof BooleanCellData) {
            BooleanCellData bcd = (BooleanCellData) cellData;
            if (bcd.isFormula()) {
                cell = row.createCell(colNum, CellType.FORMULA);
                cell.setCellValue(bcd.getValue());
                cell.setCellFormula(bcd.getFormulaValue());
            } else {
                cell = row.createCell(colNum, CellType.BOOLEAN);
                cell.setCellValue(bcd.getValue());
            }
        } else if (cellData instanceof DateCellData) {
            DateCellData drd = (DateCellData) cellData;
            if (drd.isFormula()) {
                cell = row.createCell(colNum, CellType.FORMULA);
                cell.setCellValue(drd.getValue());
                cell.setCellFormula(drd.getFormulaValue());
            } else {
                cell = row.createCell(colNum, CellType.NUMERIC);
                cell.setCellValue(drd.getValue());
            }
        } else if (cellData instanceof NumberCellData) {
            NumberCellData ncd = (NumberCellData) cellData;
            if (ncd.isFormula()) {
                cell = row.createCell(colNum, CellType.FORMULA);
                cell.setCellValue(ncd.getValue().doubleValue());
                cell.setCellFormula(ncd.getFormulaValue());
            } else {
                cell = row.createCell(colNum, CellType.NUMERIC);
                cell.setCellValue(ncd.getValue().doubleValue());
            }
        } else if (cellData instanceof StringCellData) {
            StringCellData scd = (StringCellData) cellData;
            if (scd.isFormula()) {
                cell = row.createCell(colNum, CellType.FORMULA);
                cell.setCellValue(scd.getValue());
                cell.setCellFormula(scd.getFormulaValue());
            } else {
                cell = row.createCell(colNum, CellType.STRING);
                cell.setCellValue(scd.getValue());
            }
        } else if (cellData instanceof ErrorCellData) {
            ErrorCellData ecd = (ErrorCellData) cellData;
            if (ecd.isFormula()) {
                cell = row.createCell(colNum, CellType.FORMULA);
                cell.setCellErrorValue(ecd.getValue().getCode());
                cell.setCellFormula(ecd.getFormulaValue());
            } else {
                cell = row.createCell(colNum, CellType.ERROR);
                cell.setCellErrorValue(ecd.getValue().getCode());
            }
        }

        //set style
        setCellStyle(cell, cellStyle);
    }

    private void setCellStyle(Cell cell, CellStyle cellStyle) {
        if (cell == null || cellStyle == null) {
            return;
        }

        short newStyle = cellStyle.getIndex();
        short oldStyle = cell.getCellStyle().getIndex();
        if (oldStyle != newStyle) {
            cell.setCellStyle(cellStyle);
        }
    }
}
