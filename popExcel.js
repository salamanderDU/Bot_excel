var text_log = "";
var lasrRecord;
var detNumber;
var row;
var cursor;
var importSetRef;
var parser = new sn_impex.GlideExcelParser();

var lastrow = 0;
var rowCount = 0;

var grDebtTask = new GlideRecord("x_baot_debt_sett_0_ect_debt_settlement");
grDebtTask.addEncodedQuery("sys_created_onBETWEENjavascript:gs.dateGenerate('2025-02-17','00:00:55')@javascript:gs.endOfToday()^sys_created_bySTARTSWITHBulk^number!=DEB0034393-03^ORnumber=NULL");
grDebtTask.orderBy('sys_created_on');
grDebtTask.orderBy('number');
grDebtTask.setLimit(10);
grDebtTask.query();
while (grDebtTask.next()) {
    rowCount++;
    // if (rowCount <= lastrow) {
    //     continue;
    // }

    text_log += "\n" + grDebtTask.number + " " + grDebtTask.sys_created_on;

    detNumber = grDebtTask.number;
    cursor = 1;
    createExcelParser(detNumber);


    var grImportLog = new GlideRecord("x_baot_debt_sett_0_import_excel_logs");
    grImportLog.addEncodedQuery("importsetref=" + importSetRef);
	grImportLog.orderBy('sys_created_on');
    grImportLog.orderBy('row');
    grImportLog.query();
    lastrow += grImportLog.getRowCount();
    while (grImportLog.next()) {

        row = grImportLog.row;
        lasrRecord = row;

		if(cursor == 1){
			parser.next();
		}

		if(grDebtTask.number != grImportLog.target_record){
			parser.next();
		}
		
		cursor++;
		var rowExcel = parser.getRow();
		gs.info(grImportLog.target_record+' แถวที่ ' + row + " เลขที่ยืนยันตัวตน คือ "+rowExcel["เลขที่ยืนยันตัวตน"]);
		gs.info(grDebtTask.number +" "+ grImportLog.target_record);
		if(grDebtTask.number == grImportLog.target_record){
			gs.info("Match!");
			if(cursor > lastrow){
				//do nothing
			}else{
				grDebtTask.next();
			}
			
		}
			
    }

    gs.info(text_log);
    text_log = "";

}

function createExcelParser(detNumber) {
    var grImportLog = new GlideRecord("x_baot_debt_sett_0_import_excel_logs");
    grImportLog.addEncodedQuery("target_record=" + detNumber);
    grImportLog.orderByDesc('sys_created_on');
    grImportLog.setLimit(1);
    grImportLog.query();
    if (grImportLog.next()) {
        text_log += "\n" + "ImportSetRef: " + grImportLog.importsetref;
        importSetRef = grImportLog.importsetref;

        var grIMP = new GlideRecord("x_baot_debt_sett_0_import_excel_attachments");
        grIMP.addEncodedQuery("importset_link=" + grImportLog.importsetref);
        grIMP.query();
        if (grIMP.next()) {

            text_log += "\n" + "IMP number: " + grIMP.u_number;
            text_log += " " + "IMP SYS_ID: " + grIMP.sys_id;

            var grAtt = new GlideRecord('sys_attachment');
            grAtt.addEncodedQuery("table_sys_id=" + grIMP.sys_id);
            grAtt.setLimit(1);
            grAtt.query();
            if (grAtt.next()) {

                text_log += "\nRead Excel Attachment: "+grAtt.file_name;

                var attachment = new GlideSysAttachment();
                var attachmentStream = attachment.getContentStream(grAtt.sys_id);
                parser.setSource(attachmentStream);
                parser.setHeaderRowNumber(0);
                parser.parse();
            }

        }
    }
}

