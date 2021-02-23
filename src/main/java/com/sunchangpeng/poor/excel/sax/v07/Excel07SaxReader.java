package com.sunchangpeng.poor.excel.sax.v07;

import com.sunchangpeng.poor.excel.ExcelException;
import com.sunchangpeng.poor.excel.sax.ExcelSaxReader;
import com.sunchangpeng.poor.excel.sax.ExcelSaxReaderConfig;
import com.sunchangpeng.poor.excel.utils.ExcelDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorkbookPr;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.WorkbookDocument;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;

import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;
import java.io.InputStream;
import java.util.Iterator;

public class Excel07SaxReader implements ExcelSaxReader {
    private static final String RID_PREFIX = "rId";

    @Override
    public void read(ExcelSaxReaderConfig config) {
        OPCPackage pkg;
        try {
            pkg = OPCPackage.open(config.getFile(), PackageAccess.READ);
        } catch (InvalidFormatException e) {
            throw new ExcelException(e);
        }

        read(pkg, config);
    }

    public void read(OPCPackage pkg, ExcelSaxReaderConfig config) {
        try {
            XSSFReader xssfReader = new XSSFReader(pkg);
            StylesTable stylesTable = xssfReader.getStylesTable();
            SharedStringsTable sharedStringsTable = xssfReader.getSharedStringsTable();
            boolean use1904WindowDate = use1904WindowDate(xssfReader);
            config.use1904windowing(use1904WindowDate);

            int sheetIndex = config.getTargetSheet();
            if (sheetIndex > -1) {
                InputStream inputStream = xssfReader.getSheet(RID_PREFIX + (sheetIndex + 1));
                XlsxReadContext context = new XlsxReadContext(config, sheetIndex);
                context.setStylesTable(stylesTable);
                context.setSharedStringsTable(sharedStringsTable);
                parseXmlSource(inputStream, new XlsxSaxAnalyser(context));
                config.getRowHandler().endSheet(sheetIndex);
            } else {
                final Iterator<InputStream> sheetInputStreams = xssfReader.getSheetsData();
                int index = 0;
                while (sheetInputStreams.hasNext()) {
                    XlsxReadContext context = new XlsxReadContext(config, index);
                    context.setStylesTable(stylesTable);
                    context.setSharedStringsTable(sharedStringsTable);
                    parseXmlSource(sheetInputStreams.next(), new XlsxSaxAnalyser(context));
                    config.getRowHandler().endSheet(index);
                    index++;
                }
            }
        } catch (ExcelException e) {
            throw e;
        } catch (Exception e) {
            throw new ExcelException(e);
        } finally {
            if (null != pkg) {
                try {
                    //only read
                    pkg.revert();
                } catch (Exception e) {
                    // pass
                }
            }

            ExcelDateUtil.removeCache();
        }
    }

    private boolean use1904WindowDate(XSSFReader xssfReader) throws Exception {
        InputStream workbookXml = xssfReader.getWorkbookData();
        WorkbookDocument ctWorkbook = WorkbookDocument.Factory.parse(workbookXml);
        CTWorkbook wb = ctWorkbook.getWorkbook();
        CTWorkbookPr prefix = wb.getWorkbookPr();
        return prefix != null && prefix.getDate1904();
    }

    private void parseXmlSource(InputStream inputStream, ContentHandler handler) throws Exception {
        try (final InputStream in = inputStream) {
            InputSource inputSource = new InputSource(in);
            SAXParserFactory saxFactory = SAXParserFactory.newInstance();
            saxFactory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
            saxFactory.setFeature("http://xml.org/sax/features/external-general-entities", false);
            saxFactory.setFeature("http://xml.org/sax/features/external-parameter-entities", false);
            SAXParser saxParser = saxFactory.newSAXParser();
            XMLReader xmlReader = saxParser.getXMLReader();
            xmlReader.setContentHandler(handler);
            xmlReader.parse(inputSource);
        }
    }
}
