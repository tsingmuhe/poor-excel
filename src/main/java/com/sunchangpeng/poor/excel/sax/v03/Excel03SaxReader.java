package com.sunchangpeng.poor.excel.sax.v03;

import com.sunchangpeng.poor.excel.ExcelException;
import com.sunchangpeng.poor.excel.sax.ExcelSaxReader;
import com.sunchangpeng.poor.excel.sax.ExcelSaxReaderConfig;
import com.sunchangpeng.poor.excel.utils.ExcelDateUtil;
import org.apache.poi.hssf.eventusermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.IOException;

public class Excel03SaxReader implements ExcelSaxReader {
    @Override
    public void read(ExcelSaxReaderConfig config) {
        try (POIFSFileSystem fs = new POIFSFileSystem(config.getFile())) {
            read(fs, config);
        } catch (IOException e) {
            throw new ExcelException(e);
        }
    }

    public void read(POIFSFileSystem fs, ExcelSaxReaderConfig config) throws IOException {
        XlsReadContext context = new XlsReadContext(config);

        FormatTrackingHSSFListener formatTrackingHSSFListener = new FormatTrackingHSSFListener(new MissingRecordAwareHSSFListener(new XlsSaxAnalyser(context)));
        context.setFormatTrackingHSSFListener(formatTrackingHSSFListener);

        EventWorkbookBuilder.SheetRecordCollectingListener workbookBuildingListener = new EventWorkbookBuilder.SheetRecordCollectingListener(formatTrackingHSSFListener);
        context.setHssfWorkbook(workbookBuildingListener.getStubHSSFWorkbook());

        final HSSFRequest request = new HSSFRequest();
        request.addListenerForAllRecords(formatTrackingHSSFListener);
        try {
            new HSSFEventFactory().processWorkbookEvents(request, fs);
        } finally {
            ExcelDateUtil.removeCache();
        }
    }
}
