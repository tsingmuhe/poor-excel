package com.sunchangpeng.poor.excel.utils;

import com.sunchangpeng.poor.excel.ExcelException;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.TempFile;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.util.UUID;

public class ExcelFileUtil {
    public static Workbook createBook(boolean isXlsx) {
        if (isXlsx) {
            return new XSSFWorkbook();
        }

        return new org.apache.poi.hssf.usermodel.HSSFWorkbook();
    }

    public static SXSSFWorkbook createSXSSFWorkbook(int windowSize) {
        return new SXSSFWorkbook(windowSize);
    }

    public static SXSSFWorkbook createSXSSFWorkbook(Workbook workbook, int windowSize) {
        if (workbook instanceof SXSSFWorkbook) {
            return (SXSSFWorkbook) workbook;
        }

        if (workbook instanceof XSSFWorkbook) {
            return new SXSSFWorkbook((XSSFWorkbook) workbook, windowSize);
        }

        throw new ExcelException("The input is not a [xlsx] format.");
    }

    public static boolean isXlsx(File file) {
        try {
            return FileMagic.valueOf(file) == FileMagic.OOXML;
        } catch (IOException e) {
            throw new ExcelException(e);
        }
    }

    private static String TEMP_FILE_PREFIX = System.getProperty(TempFile.JAVA_IO_TMPDIR) + File.separator + UUID.randomUUID().toString() + File.separator;
    private static String CACHE_PATH = TEMP_FILE_PREFIX + "poor-excel" + File.separator;

    public static File createPoorExcelTmpDir() throws IOException {
        return forceMkdir(new File(CACHE_PATH + UUID.randomUUID().toString()));
    }

    private static File forceMkdir(File directory) throws IOException {
        if (directory.exists()) {
            if (!directory.isDirectory()) {
                final String message = "File " + directory + " exists and is not a directory. Unable to create directory.";
                throw new IOException(message);
            }
        } else {
            if (!directory.mkdirs()) {
                // Double-check that some other thread or process hasn't made
                // the directory in the background
                if (!directory.isDirectory()) {
                    final String message = "Unable to create directory " + directory;
                    throw new IOException(message);
                }
            }
        }

        return directory;
    }
}
