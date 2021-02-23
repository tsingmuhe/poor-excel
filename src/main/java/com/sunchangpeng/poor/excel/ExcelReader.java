package com.sunchangpeng.poor.excel;

import com.sunchangpeng.poor.excel.sax.ExcelSaxReaderConfig;
import com.sunchangpeng.poor.excel.utils.ExcelFileUtil;
import org.apache.commons.io.FileUtils;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.UUID;

public interface ExcelReader {
    static ExcelSaxReaderConfig of(File file) {
        return new ExcelSaxReaderConfig(file);
    }

    /**
     *
     */
    static ExcelSaxReaderConfig of(InputStream in) throws IOException {
        File tmpDir = ExcelFileUtil.createPoorExcelTmpDir();
        File tmpFile = new File(tmpDir.getPath(), UUID.randomUUID().toString() + ".xlsx");
        FileUtils.copyToFile(in, tmpFile);
        return new ExcelSaxReaderConfig(tmpFile);
    }
}