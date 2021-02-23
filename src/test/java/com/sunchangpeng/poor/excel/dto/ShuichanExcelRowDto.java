package com.sunchangpeng.poor.excel.dto;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;
import lombok.experimental.Accessors;

import java.math.BigDecimal;
import java.util.Date;

@Getter
@Setter
@ToString
@Accessors(chain = true)
public class ShuichanExcelRowDto {
    private String pmsNo;

    private String amsNo;

    private Long poiId;

    private String poiName;

    private Long skuId;

    private String skuName;

    private BigDecimal amsHopeQuantity;

    private String containerCode;

    private BigDecimal sowingHopeQuantity;

    private BigDecimal sowingQuantity;

    private BigDecimal sowingStageQuantity;

    private Date produceDay;

    private String unitName;
}
