var text_log = "";
var lasrRecord;
var detNumber;
var row;
var cursor;
var importSetRef;
var parser = new sn_impex.GlideExcelParser();

var lastrow = 0;
var rowCount = 0;


var grDebtTask = new GlideRecord("x_baot_debt_sett_0_debt_task");
grDebtTask.addEncodedQuery("sys_created_on>javascript:gs.dateGenerate('2025-02-20','15:20:01')");//ไม่ต้อง query ใส่เป็นค่าว่าง
grDebtTask.orderBy('sys_created_on');
grDebtTask.orderBy('number');
grDebtTask.setLimit(1);
grDebtTask.query();
while (grDebtTask.next()) {


    if (grDebtTask.number == "DET0014949") {
        break;
    }

    rowCount++;
    if (rowCount <= lastrow) {
        continue;
    }

    text_log += "\n" + grDebtTask.number + " " + grDebtTask.sys_created_on;

    detNumber = grDebtTask.number;
    cursor = 2;
    createExcelParser(detNumber);


    var grImportLog = new GlideRecord("x_baot_debt_sett_0_import_excel_logs");
    grImportLog.addEncodedQuery("state=Insert^importsetref=" + importSetRef);
    grImportLog.orderBy('row');
    // grImportLog.setLimit(5);
    grImportLog.query();
    lastrow += grImportLog.getRowCount();
    while (grImportLog.next()) {

        row = grImportLog.row;
        lasrRecord = row;

        while (parseInt(row) > cursor) {
            cursor++;
            parser.next();
            if (row == cursor) {
                var rowExcel = parser.getRow();
                var MoneyValue = rowExcel["เงินต้น (บาท)"];
                gs.info(grImportLog.target_record+' เงินต้น แถวที่ ' + row + ' คือ: ' + MoneyValue+" เลขที่ยืนยันตัวตน คือ "+rowExcel["เลขที่ยืนยันตัวตน"]);
                //ใส่ค่าตรงนี้
				insertValueIdAndMoney(grImportLog.target_record, rowExcel["เลขที่ยืนยันตัวตน"], MoneyValue);
            }
        }
    }

    gs.info(text_log);
    text_log = "";

}

function createExcelParser(detNumber) {
    var grImportLog = new GlideRecord("x_baot_debt_sett_0_import_excel_logs");
    grImportLog.addEncodedQuery("state=Insert^target_record=" + detNumber);
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
                parser.setHeaderRowNumber(1);
                parser.parse();
            }

        }
    }
}

function insertValueIdAndMoney(detNumber, iden, money) {
   var grDebtTask = new GlideRecord("x_baot_debt_sett_0_debt_task");
   grDebtTask.addQuery('number',detNumber);
   grDebtTask.query();
    if(grDebtTask.next()){
		grDebtTask.u_principle = set_money_object(money); //เงินต้น
		grDebtTask.u_walkin_identification_number = iden; //เลขที่ยืนยันตัวตน
		grDebtTask.update();
		gs.info(detNumber+" อัปเดตสำเร็จ duke");
    }
}

function set_money_object(money_input) {
	var x = money_input.replace("฿", "");
	var y = x.replace(",", "");
	return "THB;" + y;
}