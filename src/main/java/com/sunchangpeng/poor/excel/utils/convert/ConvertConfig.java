package com.sunchangpeng.poor.excel.utils.convert;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;
import lombok.experimental.Accessors;

@Getter
@Setter
@ToString
@Accessors(chain = true)
public class ConvertConfig {
    private String datePattern;
    private boolean use1904windowing;
}
