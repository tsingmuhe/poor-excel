package com.sunchangpeng.poor.excel;

import com.alibaba.fastjson.JSON;
import com.sunchangpeng.poor.excel.cell.CellData;
import com.sunchangpeng.poor.excel.dto.ShuichanExcelRowDto;
import com.sunchangpeng.poor.excel.sax.RowHandler;
import com.sunchangpeng.poor.excel.utils.convert.ConvertConfig;
import org.apache.commons.collections4.MapUtils;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import static com.sunchangpeng.poor.excel.utils.ExcelConvertUtil.*;

public class ExcelReaderTest {
    @Test
    public void test() {
        ZipSecureFile.setMinInflateRatio(-1.0d);

        List<ShuichanExcelRowDto> rowDtos = new ArrayList<>();
        ExcelReader.of(new File("/Users/sunchangpeng/workspace/java/poor-excel/src/test/resources/测试.xlsx"))
                .skipEmptyRow(true)
                .rowFilter((sheetNum, rowNum) -> {
                    if (sheetNum == 0) {
                        if (rowNum <= 0) {
                            return false;
                        }

                        return true;
                    }

                    if (sheetNum == 1 && rowNum == 0) {
                        return true;
                    }
                    return false;
                })
                .rowHandler((sheetIndex, rowIndex, cellMap) -> {
                    if (MapUtils.isEmpty(cellMap)) {
                        return;
                    }

                    if (sheetIndex == 1) {
                        System.out.println(getString(cellMap.get(0)));
                        return;
                    }

                    rowDtos.add(new ShuichanExcelRowDto()
                            .setPmsNo(getString(cellMap.get(0)))
                            .setAmsNo(getString(cellMap.get(1)))
                            .setPoiId(getLong(cellMap.get(2)))
                            .setPoiName(getString(cellMap.get(3)))
                            .setSkuId(getLong(cellMap.get(4)))
                            .setSkuName(getString(cellMap.get(5)))
                            .setAmsHopeQuantity(getBigDecimal(cellMap.get(6)))
                            .setContainerCode(getString(cellMap.get(7)))
                            .setSowingHopeQuantity(getBigDecimal(cellMap.get(8)))
                            .setSowingQuantity(getBigDecimal(cellMap.get(9)))
                            .setSowingStageQuantity(getBigDecimal(cellMap.get(10)))
                            .setProduceDay(getDate(cellMap.get(11), new ConvertConfig().setDatePattern("yyyy/MM/dd")))
                            .setUnitName(getString(cellMap.get(12))));
                })
                .read();
        System.out.println(JSON.toJSONString(rowDtos));
    }

    @Test
    public void test_2() throws IOException {
        ZipSecureFile.setMinInflateRatio(-1.0d);
        ExcelReader.of(new FileInputStream("/Users/sunchangpeng/workspace/java/poor-excel/src/test/resources/测试.xlsx"))
                .rowHandler(new RowHandler() {
                    @Override
                    public void handle(int sheetIndex, long rowIndex, Map<Integer, CellData> cellMap) {
                        System.out.println(cellMap);
                    }
                }).read();
    }
}