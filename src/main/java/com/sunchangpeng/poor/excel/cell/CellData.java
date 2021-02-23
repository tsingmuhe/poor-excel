package com.sunchangpeng.poor.excel.cell;

import com.sunchangpeng.poor.excel.write.style.CellDataStyle;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;
import lombok.experimental.Accessors;

@Getter
@Setter
@ToString
@Accessors(chain = true)
public abstract class CellData<T> {
    private boolean formula;
    private String formulaValue;

    private T value;

    /**
     * for write
     */
    private CellDataStyle style;
}
