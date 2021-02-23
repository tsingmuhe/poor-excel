package com.sunchangpeng.poor.excel.write.style;

import lombok.Getter;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;

import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;
import java.util.Optional;

public class ExcelWriteStyleSet {
    private final Workbook workbook;
    private final Map<CellDataStyle, CellStyle> styleMap = new HashMap<>();
    @Getter
    private CellStyle globalCellStyle;

    public ExcelWriteStyleSet(Workbook workbook) {
        this.workbook = workbook;
        initCellStyle();
    }

    public void resetGlobalCellStyle(CellDataStyle styleCss) {
        this.globalCellStyle = getOrCreateCellStyle(styleCss);
    }

    public CellStyle getOrCreateCellStyle(CellDataStyle styleCss) {
        if (styleCss == null) {
            return null;
        }

        CellStyle result = this.styleMap.get(styleCss);
        if (result != null) {
            return result;
        }

        CellStyle cellStyle = convertToCellStyle(styleCss);
        this.styleMap.put(styleCss, cellStyle);
        return cellStyle;
    }

    public Font createFont() {
        return this.workbook.createFont();
    }

    private void initCellStyle() {
        int numberCellStyles = this.workbook.getNumCellStyles();

        for (int i = 0; i < numberCellStyles; i++) {
            CellStyle wbStyle = workbook.getCellStyleAt(i);
            CellDataStyle styleCss = convertToCellDataStyle(wbStyle);
            this.styleMap.put(styleCss, wbStyle);
        }
    }

    private CellStyle convertToCellStyle(CellDataStyle style) {
        CellStyle result = this.workbook.createCellStyle();

        if (StringUtils.isNotBlank(style.getDataFormat())) {
            result.setDataFormat(workbook.createDataFormat().getFormat(style.getDataFormat()));
        }

        if (style.getFont() != null) {
            result.setFont(workbook.getFontAt(style.getFont()));
        }

        if (StringUtils.isNotBlank(style.getTextAlign())) {
            Optional<HorizontalAlignment> align = Arrays.stream(HorizontalAlignment.values())
                    .filter(item -> item.name().toLowerCase().equals(style.getTextAlign())).findFirst();
            if (align.isPresent()) {
                result.setAlignment(align.get());
            }
        }

        if (StringUtils.isNotBlank(style.getTextVerticalAlign())) {
            Optional<VerticalAlignment> align = Arrays.stream(VerticalAlignment.values())
                    .filter(item -> item.name().toLowerCase().equals(style.getTextVerticalAlign())).findFirst();
            if (align.isPresent()) {
                result.setVerticalAlignment(align.get());
            }
        }

        if (style.getTextWrap() != null) {
            result.setWrapText(style.getTextWrap());
        }

        if (style.getTextIndent() != null) {
            result.setIndention(style.getTextIndent());
        }

        if (StringUtils.isNotBlank(style.getTopBorderStyle())) {
            Optional<BorderStyle> borderStyle = Arrays.stream(BorderStyle.values())
                    .filter(item -> item.name().toLowerCase().equals(style.getTopBorderStyle())).findFirst();
            if (borderStyle.isPresent()) {
                result.setBorderTop(borderStyle.get());
            }
        }

        if (StringUtils.isNotBlank(style.getRightBorderStyle())) {
            Optional<BorderStyle> borderStyle = Arrays.stream(BorderStyle.values())
                    .filter(item -> item.name().toLowerCase().equals(style.getRightBorderStyle())).findFirst();
            if (borderStyle.isPresent()) {
                result.setBorderRight(borderStyle.get());
            }
        }

        if (StringUtils.isNotBlank(style.getBottomBorderStyle())) {
            Optional<BorderStyle> borderStyle = Arrays.stream(BorderStyle.values())
                    .filter(item -> item.name().toLowerCase().equals(style.getBottomBorderStyle())).findFirst();
            if (borderStyle.isPresent()) {
                result.setBorderBottom(borderStyle.get());
            }
        }

        if (StringUtils.isNotBlank(style.getLeftBorderStyle())) {
            Optional<BorderStyle> borderStyle = Arrays.stream(BorderStyle.values())
                    .filter(item -> item.name().toLowerCase().equals(style.getLeftBorderStyle())).findFirst();
            if (borderStyle.isPresent()) {
                result.setBorderLeft(borderStyle.get());
            }
        }

        if (style.getTopBorderColor() != null) {
            result.setTopBorderColor(style.getTopBorderColor());
        }

        if (style.getRightBorderColor() != null) {
            result.setRightBorderColor(style.getRightBorderColor());
        }

        if (style.getBottomBorderColor() != null) {
            result.setBottomBorderColor(style.getBottomBorderColor());
        }

        if (style.getLeftBorderColor() != null) {
            result.setLeftBorderColor(style.getLeftBorderColor());
        }

        if (StringUtils.isNotBlank(style.getFillPattern())) {
            Optional<FillPatternType> fp = Arrays.stream(FillPatternType.values())
                    .filter(item -> item.name().toLowerCase().equals(style.getFillPattern())).findFirst();
            if (fp.isPresent()) {
                result.setFillPattern(fp.get());
            }
        }

        if (style.getBackgroundColor() != null) {
            result.setFillBackgroundColor(style.getBackgroundColor());
        }

        if (style.getForegroundColor() != null) {
            result.setFillForegroundColor(style.getForegroundColor());
        }

        if (style.getHidden() != null) {
            result.setHidden(style.getHidden());
        }

        if (style.getLocked() != null) {
            result.setLocked(style.getLocked());
        }

        if (style.getQuotePrefix() != null) {
            result.setQuotePrefixed(style.getQuotePrefix());
        }

        if (style.getRotation() != null) {
            result.setRotation(style.getRotation());
        }

        if (style.getShrink() != null) {
            result.setShrinkToFit(style.getShrink());
        }

        return result;
    }

    private CellDataStyle convertToCellDataStyle(CellStyle style) {
        return new CellDataStyle()
                //
                .setDataFormat(style.getDataFormatString())
                //
                .setFont(style.getFontIndexAsInt())
                .setTextAlign(StringUtils.lowerCase(style.getAlignment().name()))
                .setTextVerticalAlign(StringUtils.lowerCase(style.getVerticalAlignment().name()))
                .setTextWrap(style.getWrapText())
                .setTextIndent(style.getIndention())
                //
                .setTopBorderStyle(StringUtils.lowerCase(style.getBorderTop().name()))
                .setRightBorderStyle(StringUtils.lowerCase(style.getBorderRight().name()))
                .setBottomBorderStyle(StringUtils.lowerCase(style.getBorderBottom().name()))
                .setLeftBorderStyle(StringUtils.lowerCase(style.getBorderLeft().name()))
                //
                .setTopBorderColor(style.getTopBorderColor())
                .setRightBorderColor(style.getRightBorderColor())
                .setBottomBorderColor(style.getBottomBorderColor())
                .setLeftBorderColor(style.getLeftBorderColor())
                //
                .setFillPattern(StringUtils.lowerCase(style.getFillPattern().name()))
                .setForegroundColor(style.getFillForegroundColor())
                .setBackgroundColor(style.getFillBackgroundColor())
                //
                .setHidden(style.getHidden())
                .setLocked(style.getLocked())
                .setQuotePrefix(style.getQuotePrefixed())
                .setRotation(style.getRotation())
                .setShrink(style.getShrinkToFit());
    }
}
