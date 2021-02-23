package com.sunchangpeng.poor.excel.write.style;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;
import lombok.experimental.Accessors;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.util.Objects;

@Getter
@Setter
@ToString
@Accessors(chain = true)
public class CellDataStyle {
    //------------------number format
    private String dataFormat;

    //------------------text
    private Integer font;

    /**
     * general
     * left
     * center
     * right
     * fill
     * justify
     * center_selection
     * distributed
     */
    private String textAlign;

    /**
     * top
     * center
     * bottom
     * justify
     * distributed
     */
    private String textVerticalAlign;
    private Boolean textWrap;
    private Short textIndent;

    //------------------border
    /**
     * none
     * thin
     * medium
     * dashed
     * dotted
     * thick
     * double
     * hair
     * medium_dashed
     * dash_dot
     * medium_dash_dot
     * dash_dot_dot
     * medium_dash_dot_dot
     * slanted_dash_dot
     */
    private String topBorderStyle;
    private String rightBorderStyle;
    private String bottomBorderStyle;
    private String leftBorderStyle;

    /**
     * @see IndexedColors
     */
    private Short topBorderColor;
    private Short rightBorderColor;
    private Short bottomBorderColor;
    private Short leftBorderColor;

    //------------------groundColor
    /**
     * no_fill
     * solid_foreground
     * fine_dots
     * alt_bars
     * sparse_dots
     * thick_horz_bands
     * thick_vert_bands
     * thick_backward_diag
     * thick_forward_diag
     * big_spots
     * bricks
     * thin_horz_bands
     * thin_vert_bands
     * thin_backward_diag
     * thin_forward_diag
     * squares
     * diamonds
     * less_dots
     * least_dots
     */
    private String fillPattern;
    private Short foregroundColor;
    private Short backgroundColor;

    //------------------else
    private Boolean hidden;
    private Boolean locked;

    private Boolean quotePrefix;
    private Short rotation;
    private Boolean shrink;

    public CellDataStyle setBorderStyle(String borderStyle) {
        this.topBorderStyle = borderStyle;
        this.rightBorderStyle = borderStyle;
        this.bottomBorderStyle = borderStyle;
        this.leftBorderStyle = borderStyle;
        return this;
    }

    public CellDataStyle setBorderColor(Short borderColor) {
        this.topBorderColor = borderColor;
        this.rightBorderColor = borderColor;
        this.bottomBorderColor = borderColor;
        this.leftBorderColor = borderColor;
        return this;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) {
            return true;
        }
        if (o == null || getClass() != o.getClass()) {
            return false;
        }
        CellDataStyle styleCss = (CellDataStyle) o;
        return Objects.equals(dataFormat, styleCss.dataFormat)
                && Objects.equals(font, styleCss.font)
                && Objects.equals(textAlign, styleCss.textAlign)
                && Objects.equals(textVerticalAlign, styleCss.textVerticalAlign)
                && Objects.equals(textWrap, styleCss.textWrap)
                && Objects.equals(textIndent, styleCss.textIndent)
                && Objects.equals(topBorderStyle, styleCss.topBorderStyle)
                && Objects.equals(rightBorderStyle, styleCss.rightBorderStyle)
                && Objects.equals(bottomBorderStyle, styleCss.bottomBorderStyle)
                && Objects.equals(leftBorderStyle, styleCss.leftBorderStyle)
                && Objects.equals(topBorderColor, styleCss.topBorderColor)
                && Objects.equals(rightBorderColor, styleCss.rightBorderColor)
                && Objects.equals(bottomBorderColor, styleCss.bottomBorderColor)
                && Objects.equals(leftBorderColor, styleCss.leftBorderColor)
                && Objects.equals(fillPattern, styleCss.fillPattern)
                && Objects.equals(foregroundColor, styleCss.foregroundColor)
                && Objects.equals(backgroundColor, styleCss.backgroundColor)
                && Objects.equals(hidden, styleCss.hidden)
                && Objects.equals(locked, styleCss.locked)
                && Objects.equals(quotePrefix, styleCss.quotePrefix)
                && Objects.equals(rotation, styleCss.rotation)
                && Objects.equals(shrink, styleCss.shrink);
    }

    @Override
    public int hashCode() {
        return Objects.hash(dataFormat,
                font,
                textAlign,
                textVerticalAlign,
                textWrap,
                textIndent,
                topBorderStyle,
                rightBorderStyle,
                bottomBorderStyle,
                leftBorderStyle,
                topBorderColor,
                rightBorderColor,
                bottomBorderColor,
                leftBorderColor,
                fillPattern,
                foregroundColor,
                backgroundColor,
                hidden,
                locked,
                quotePrefix,
                rotation,
                shrink);
    }

    public CellDataStyle cloneStyle() {
        return new CellDataStyle()
                .setDataFormat(this.getDataFormat())
                .setFont(this.getFont())
                .setTextAlign(this.getTextAlign())
                .setTextVerticalAlign(this.getTextVerticalAlign())
                .setTextWrap(this.getTextWrap())
                .setTextIndent(this.getTextIndent())
                .setTopBorderStyle(this.getTopBorderStyle())
                .setRightBorderStyle(this.getRightBorderStyle())
                .setBottomBorderStyle(this.getBottomBorderStyle())
                .setLeftBorderStyle(this.getLeftBorderStyle())
                .setTopBorderColor(this.getTopBorderColor())
                .setRightBorderColor(this.getRightBorderColor())
                .setBottomBorderColor(this.getBottomBorderColor())
                .setLeftBorderColor(this.getLeftBorderColor())
                .setFillPattern(this.getFillPattern())
                .setForegroundColor(this.getForegroundColor())
                .setBackgroundColor(this.getBackgroundColor())
                .setHidden(this.getHidden())
                .setLocked(this.getLocked())
                .setQuotePrefix(this.getQuotePrefix())
                .setRotation(this.getRotation())
                .setShrink(this.getShrink());
    }
}
