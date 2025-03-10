var text_log = "";
var lasrRecord;
var detNumber;
var row;
var cursor;
var importSetRef;
var parser = new sn_impex.GlideExcelParser();

var lastrow = 0;
var rowCount = 1;
var limitRow = 88;

var grDebtTask = new GlideRecord("x_baot_debt_sett_0_ect_debt_settlement");
grDebtTask.addEncodedQuery("u_resolution_code=200^ORu_resolution_code=100^u_glideline_debtISEMPTY^u_reason_unable_helpISEMPTY^sys_created_on<javascript:gs.dateGenerate('2025-02-23','00:00:31')^state=200");
// grDebtTask.orderBy('sys_created_on');
grDebtTask.orderBy('number');
grDebtTask.setLimit(limitRow);
grDebtTask.query();
while (rowCount < limitRow) {

	gs.info("rowCount: "+rowCount);
	if(rowCount == 1){
		grDebtTask.next();
	}

    text_log += "\n" + grDebtTask.number + " " + grDebtTask.sys_created_on;
    detNumber = grDebtTask.number;
    cursor = 1;
    createExcelParser(detNumber);

    var grImportLog = new GlideRecord("x_baot_debt_sett_0_import_excel_logs");
    grImportLog.addEncodedQuery("importsetref=" + importSetRef);
    grImportLog.orderBy('row');
    grImportLog.query();
    lastrow += grImportLog.getRowCount();
	var completeFix = "No";
	var preV = 0;
    while (grImportLog.next()) {
		///////
		if(parseInt(preV) !== parseInt(grImportLog.row)){
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

		if(grDebtTask.number == grImportLog.target_record){
			gs.info("Debt set: "+grDebtTask.number +" log Target: "+ grImportLog.target_record);
			gs.info(cursor);

			if(grDebtTask.number == rowExcel["หมายเลข Case Task"]){
				gs.info("number equal!");
			}else{
				gs.info("number diff!");
			}
			gs.info(grDebtTask.number+" หมายเลข Case Task "+rowExcel["หมายเลข Case Task"]+" "+row+" ผลการพิจารณา: "+rowExcel["ผลการพิจารณา"]);
			gs.info(grImportLog.target_record+' แถวที่ ' + row + " แนวทางกับเหตุผล คือ "+rowExcel["แนวทางการช่วยเหลือลูกหนี้"]+" "+rowExcel["เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้"]);
		
			gs.info("Match!");
			grDebtTask.next();
			rowCount++;
			completeFix = "yes";
		}
		preV = parseInt(grImportLog.row);
		}
		/////
    }

	if(completeFix === "No"){

		gs.info("[Error] "+grDebtTask.number+" is can't fix ");

		grDebtTask.next();
		
	}

    gs.info(text_log);
    text_log = "";

}

function createExcelParser(detNumber) {
    var grImportLog = new GlideRecord("x_baot_debt_sett_0_import_excel_logs");
    grImportLog.addEncodedQuery("target_record=" + detNumber+"^message=สำเร็จ");
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

				// gs.info(JSON.stringify(parser.getColumnHeaders()));
            }

        }
    }
}

