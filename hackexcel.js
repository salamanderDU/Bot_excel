var text_log = "";
var grDebtTask = new GlideRecord("x_baot_debt_sett_0_debt_task");
grDebtTask.addEncodedQuery("sys_created_on>javascript:gs.dateGenerate('2025-02-14','15:25:01')");
grDebtTask.orderBy('sys_created_on');
grDebtTask.setLimit(20);
grDebtTask.query();
// loop ดูข้อมูลทั้งหมดของ record ที่อยู่ใน table นี้
while(grDebtTask.next()){
	//ใส่ว่าควรจะหยุดตอนไหน
	if(grDebtTask.number == "DET0014949"){
		break;
	}
	
	text_log += "\n"+grDebtTask.number+" "+grDebtTask.sys_created_on;
	// เดี๋ยวหา flag มาใส่เพิ่มให้มัน continue

	// ไปเอา importsetref ที่อยู่ในlog ที่ insert สำเร็จมาดู
	var grImportLog = new GlideRecord("x_baot_debt_sett_0_import_excel_logs");
	grImportLog.addEncodedQuery("state=Insert^target_record="+ grDebtTask.number);
	grImportLog.orderByDesc('sys_created_on');
	grImportLog.setLimit(1);
	grImportLog.query();
	if(grImportLog.next()){
		text_log += "\n"+"ImportSetRed: "+grImportLog.importsetref;
		var grIMP = new GlideRecord("x_baot_debt_sett_0_import_excel_attachments");
		grIMP.addEncodedQuery("importset_link="+grImportLog.importsetref);
		grIMP.query();
		if(grIMP.next()){
			text_log += "\n"+"IMP number: "+grIMP.u_number;
			text_log += "\n"+"IMP SYS_ID: "+grIMP.sys_id;
			//เข้าไปเอา excel มาเทียบ
			var grAtt = new GlideRecord('sys_attachment');
			grAtt.addEncodedQuery("table_sys_id=" + grIMP.sys_id);
			grAtt.setLimit(1);
			grAtt.query();
			if(grAtt.next()){
				text_log += "\nget att";
				var parser = new sn_impex.GlideExcelParser();
				var attachment = new GlideSysAttachment();
				var attachmentStream = attachment.getContentStream(grAtt.sys_id);
				parser.setSource(attachmentStream);
				var list_sheet_name = parser.getSheetNames();
				parser.setSheetName(list_sheet_name[0]); // Select First Sheet Excel
				if (parser.parse()) {
					var headers = parser.getColumnHeaders(); // retrieve the column headers
					if (parser.next()) {
						var header_row2 = Object.values(parser.getRow());
						gs.info(JSON.stringify(header_row2));
					}
				}
			}

		}
	}
	gs.info(text_log);
	text_log = "";
}