package com.sunchangpeng.poor.excel.sax.v03;

import com.sunchangpeng.poor.excel.sax.v03.handlers.*;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.record.*;

import java.util.HashMap;
import java.util.Map;

public class XlsSaxAnalyser implements HSSFListener {
    private static final short DUMMY_RECORD_SID = -1;

    private static final Map<Short, XlsRecordHandler> XLS_RECORD_HANDLER_MAP = new HashMap<>(16);

    static {
        //cell
        XLS_RECORD_HANDLER_MAP.put(BoolErrRecord.sid, new BoolErrRecordHandler());
        XLS_RECORD_HANDLER_MAP.put(LabelRecord.sid, new LabelRecordHandler());
        XLS_RECORD_HANDLER_MAP.put(LabelSSTRecord.sid, new LabelSstRecordHandler());
        XLS_RECORD_HANDLER_MAP.put(NumberRecord.sid, new NumberRecordHandler());
        XLS_RECORD_HANDLER_MAP.put(FormulaRecord.sid, new FormulaRecordHandler());
        XLS_RECORD_HANDLER_MAP.put(StringRecord.sid, new StringRecordHandler());

        XLS_RECORD_HANDLER_MAP.put(DUMMY_RECORD_SID, new DummyRecordHandler());

        //sheet
        XLS_RECORD_HANDLER_MAP.put(BOFRecord.sid, new BofRecordHandler());
        XLS_RECORD_HANDLER_MAP.put(EOFRecord.sid, new EofRecordHandler());

        //else
        XLS_RECORD_HANDLER_MAP.put(BoundSheetRecord.sid, new BoundSheetRecordHandler());
        XLS_RECORD_HANDLER_MAP.put(SSTRecord.sid, new SstRecordHandler());
    }

    private final XlsReadContext context;

    public XlsSaxAnalyser(XlsReadContext context) {
        this.context = context;
    }

    @Override
    public void processRecord(Record record) {
        XlsRecordHandler handler = XLS_RECORD_HANDLER_MAP.get(record.getSid());
        if (handler == null) {
            return;
        }

        if (this.context.getConfig().getTargetSheet() > -1 && this.context.getCurrentSheetIndex() > this.context.getConfig().getTargetSheet()) {
            return;
        }

        if (!(handler instanceof IgnorableXlsRecordHandler) || isProcessCurrentSheet()) {
            //need to read the current sheet
            handler.processRecord(context, record);
        }
    }

    private boolean isProcessCurrentSheet() {
        return this.context.getConfig().getTargetSheet() < 0 || this.context.getCurrentSheetIndex() == this.context.getConfig().getTargetSheet();
    }
}
