package com.sunchangpeng.poor.excel.sax.v07.tag;

import com.sunchangpeng.poor.excel.ExcelException;
import com.sunchangpeng.poor.excel.cell.*;
import com.sunchangpeng.poor.excel.sax.v07.XlsxCellDataType;
import com.sunchangpeng.poor.excel.sax.v07.XlsxCellRuntime;
import com.sunchangpeng.poor.excel.sax.v07.XlsxReadContext;
import com.sunchangpeng.poor.excel.utils.ExcelDateUtil;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.xml.sax.Attributes;

import java.math.BigDecimal;
import java.util.Map;

import static org.apache.poi.ss.usermodel.DateUtil.getJavaDate;

public class CellTagHandler implements IgnorableXlsxTagHandler {
    @Override
    public void startElement(XlsxReadContext context, String name, Attributes attributes) {
        String rValue = AttributeName.r.getValue(attributes);
        if (rValue == null) {
            return;
        }

        int currentColumnIndex = getCol(rValue);

        context.setLastColumnIndex(currentColumnIndex);
        context.setCellRuntime(initXlsxCellRuntime(context.getStylesTable(), attributes));
    }

    private int getCol(String rValue) {
        int col = 0;
        char[] currentIndex = rValue.replaceAll("[0-9]", "").toCharArray();
        for (int i = 0; i < currentIndex.length; i++) {
            col += (currentIndex[i] - '@') * Math.pow(26, (currentIndex.length - i - 1));
        }
        return col - 1;
    }

    private XlsxCellRuntime initXlsxCellRuntime(StylesTable stylesTable, Attributes attributes) {
        XlsxCellRuntime cellRuntime = new XlsxCellRuntime();

        cellRuntime.setCellDataType(XlsxCellDataType.parseCellDataType(AttributeName.t.getValue(attributes)));

        String sValue = AttributeName.s.getValue(attributes);

        // style.xml->cellXfs->xf
        int xfIndex = 0;
        if (null != sValue) {
            xfIndex = Integer.parseInt(sValue);
        }

        XSSFCellStyle xssfCellStyle = stylesTable.getStyleAt(xfIndex);
        int dataFormat = xssfCellStyle.getDataFormat();
        String dataFormatString = BuiltinFormats.getBuiltinFormat(dataFormat);
        if (dataFormatString == null) {
            dataFormatString = xssfCellStyle.getDataFormatString();
        }

        cellRuntime.setDataFormat(dataFormat);
        cellRuntime.setDataFormatString(dataFormatString);
        cellRuntime.setOriginValue(new StringBuilder());

        return cellRuntime;
    }

    @Override
    public void endElement(XlsxReadContext context, String name) {
        XlsxCellRuntime cellRuntime = context.getCellRuntime();
        checkEmpty(cellRuntime);

        Map<Integer, CellData> cellMap = context.getCellMap();
        XlsxCellDataType cellDataType = cellRuntime.getCellDataType();
        switch (cellDataType) {
            case STRING:
            case RICH_TEXT_STRING:
            case DIRECT_STRING:
                cellMap.put(context.getLastColumnIndex(), new StringCellData()
                        .setFormula(cellRuntime.isFormula()).setFormulaValue(cellRuntime.getFormulaValue())
                        .setValue(cellRuntime.getValue()));
                break;
            case ERROR:
                cellMap.put(context.getLastColumnIndex(), new ErrorCellData()
                        .setFormula(cellRuntime.isFormula()).setFormulaValue(cellRuntime.getFormulaValue())
                        .setValue(FormulaError.forString(cellRuntime.getValue())));
                break;
            case BOOLEAN:
                cellMap.put(context.getLastColumnIndex(), new BooleanCellData()
                        .setFormula(cellRuntime.isFormula()).setFormulaValue(cellRuntime.getFormulaValue())
                        .setValue("1".equals(cellRuntime.getValue())));
                break;
            case NUMBER:
                cellMap.put(context.getLastColumnIndex(), dealNumberCell(context, cellRuntime));
                break;
            case EMPTY:
                break;
            default:
                throw new ExcelException("Illegal XlsxCellDataType");
        }

        context.setLastColumnIndex(-1);
        context.setCellRuntime(null);
    }

    private void checkEmpty(XlsxCellRuntime cellRuntime) {
        XlsxCellDataType type = cellRuntime.getCellDataType();
        if (null == type || XlsxCellDataType.EMPTY == type) {
            return;
        }

        String str = cellRuntime.getValue();
        if (str == null || "".equals(str)) {
            if (!cellRuntime.isFormula()) {
                cellRuntime.setCellDataType(XlsxCellDataType.EMPTY);
            }
        }
    }

    private CellData dealNumberCell(XlsxReadContext context, XlsxCellRuntime cellRuntime) {
        if (ExcelDateUtil.isADateFormat(cellRuntime.getDataFormat(), cellRuntime.getDataFormatString())) {
            return new DateCellData()
                    .setFormula(cellRuntime.isFormula()).setFormulaValue(cellRuntime.getFormulaValue())
                    .setValue(getJavaDate(Double.parseDouble(cellRuntime.getValue()), context.getConfig().isUse1904windowing()));
        }

        return new NumberCellData().setFormula(cellRuntime.isFormula()).setFormulaValue(cellRuntime.getFormulaValue())
                .setValue(new BigDecimal(cellRuntime.getValue()));
    }
}
