package com.sunchangpeng.poor.excel;

import com.sunchangpeng.poor.excel.cell.*;
import com.sunchangpeng.poor.excel.write.ExcelWriteExecutor;
import com.sunchangpeng.poor.excel.write.style.CellDataStyle;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.junit.Test;

import java.io.File;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class ExcelWriterTest {
    @Test
    public void test() {
        ExcelWriteExecutor writeExecutor = ExcelWriter.xlsx();
        writeExecutor.globalStyle(createHeadCellStyle());

        List<CellData> data = new ArrayList<>();

        CellData cell2 = new BooleanCellData().setValue(true);
        CellData cell2_1 = new BooleanCellData().setValue(true).setFormula(true).setFormulaValue("B1");

        CellData cell3 = new DateCellData().setValue(new Date());
        CellData cell3_1 = new DateCellData().setValue(new Date()).setFormula(true).setFormulaValue("D1");

        CellData cell4 = new ErrorCellData().setValue(FormulaError.VALUE);
        CellData cell4_1 = new ErrorCellData().setValue(FormulaError.VALUE).setFormula(true).setFormulaValue("F1");

        CellData cell5 = new NumberCellData().setValue(new BigDecimal("3.1415926"));
        CellData cell5_1 = new NumberCellData().setValue(new BigDecimal("3.1415926")).setFormula(true).setFormulaValue("A5");

        CellData cell6 = new StringCellData().setValue("sunchp");
        CellData cell6_1 = new StringCellData().setValue("sunchp").setFormula(true).setFormulaValue("J1");

        data.add(null);

        data.add(cell2);
        data.add(cell2_1);

        data.add(cell3);
        data.add(cell3_1);

        data.add(cell4);
        data.add(cell4_1);

        data.add(cell5);
        data.add(cell5_1);

        data.add(cell6);
        data.add(cell6_1);

        writeExecutor.appendRow(data);
        try {
            writeExecutor.flush(new File("/Users/sunchangpeng/Downloads/poor_excel.xlsx")).close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public CellDataStyle createHeadCellStyle() {
        final CellDataStyle cellStyle = new CellDataStyle();
        cellStyle.setTextAlign("center");
        cellStyle.setTextVerticalAlign("center");
        cellStyle.setBorderStyle("thin");
        cellStyle.setBorderColor(IndexedColors.RED.index);
        cellStyle.setFillPattern("solid_foreground");
        cellStyle.setForegroundColor(IndexedColors.GREY_25_PERCENT.index);
        return cellStyle;
    }
}