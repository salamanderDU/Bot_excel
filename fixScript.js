var text_log = "";
var lasrRecord;
var detNumber;
var row;
var cursor;
var importSetRef;
var parser = new sn_impex.GlideExcelParser();

var lastrow = 0;
var rowCount = 1;
var limitRow = 100;

var allRecordSysId = [];
var notMatchNumber = [];
var notMatchLog = [];

var grDebtTask = new GlideRecord("x_baot_debt_sett_0_ect_debt_settlement");
grDebtTask.addEncodedQuery("state=3^u_resolution_code=100^ORu_resolution_code=200^u_glideline_debtISEMPTY^u_reason_unable_helpISEMPTY^sys_created_byNOT LIKEbulk^sys_class_name=x_baot_debt_sett_0_ect_debt_settlement");
// grDebtTask.orderBy('sys_created_on');
grDebtTask.orderBy('number');
grDebtTask.setLimit(limitRow);
grDebtTask.query();
while (rowCount < limitRow) {

    gs.info("rowCount: " + rowCount);
    if (rowCount == 1) {
        grDebtTask.next();
		allRecordSysId.push(grDebtTask.sys_id.toString());
    }

	gs.info("SYSID: "+grDebtTask.sys_id.toString());

    

    text_log += "\n" + grDebtTask.number + " " + grDebtTask.sys_created_on;
    detNumber = grDebtTask.number;
    cursor = 1;
    if(createExcelParser(detNumber)){
		notMatchLog.push(grDebtTask.sys_id.toString());
		grDebtTask.next();
		continue;
	}

    var grImportLog = new GlideRecord("x_baot_debt_sett_0_import_excel_logs");
    grImportLog.addEncodedQuery("importsetref=" + importSetRef);
    grImportLog.orderBy('row');
    grImportLog.query();
    lastrow += grImportLog.getRowCount();
    var preV = 0;
    while (grImportLog.next()) {

        if (parseInt(preV) !== parseInt(grImportLog.row)) {
            row = grImportLog.row;
            lasrRecord = row;

            if (cursor == 1) {
                parser.next();
            }

            if (grDebtTask.number != grImportLog.target_record) {
                parser.next();
            }

            cursor++;

            var rowExcel = parser.getRow();

            if (grDebtTask.number == grImportLog.target_record) {
                gs.info("cursor: " + cursor);

                if (grDebtTask.number == rowExcel["หมายเลข Case Task"]) {
                    gs.info("number equal! " + grDebtTask.number + " " + rowExcel["หมายเลข Case Task"]);
					var simPercent = 40;
					if(compareStrings(grDebtTask.u_provider.getDisplayValue(),rowExcel["ผู้ให้บริการ"]),simPercent){

                        if (rowExcel["ผลการพิจารณา"] == "ได้ข้อสรุปกับลูกค้า" && grDebtTask.u_resolution_code == "100") {
                            var glideline = rowExcel["แนวทางการช่วยเหลือลูกหนี้"].split("|")[0].trim();
                            var glideline_debt = get_select_valid_active_check_fair('u_glideline_debtor', 'u_name', glideline, grDebtTask.u_debt_project.getDisplayValue());
                            if (glideline_debt == "Error") {
                                gs.info("[Error]] ค่า แนวทางการช่วยเหลือลูกหนี้ ไม่ถูกต้อง");
                            } else {
                                gs.info(grDebtTask.number + " [Update] แนวทางการช่วยเหลือลูกหนี้ => [" + glideline + "] " + glideline_debt);
                                // grDebtTask.u_glideline_debt = glideline_debt;
                                // grDebtTask.update();
                            }

                        } else if (rowExcel["ผลการพิจารณา"] == "ไม่ได้ข้อสรุปกับลูกค้า" && grDebtTask.u_resolution_code == "200") {
                            var reason = rowExcel["เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้"].split("|")[0].trim();
                            var reason_unable_help = get_select_valid_active_check_fair('u_unable_help', 'u_name', reason, grDebtTask.u_debt_project.getDisplayValue());
                            if (reason_unable_help == "Error") {
                                gs.info("[Error] ค่า เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้ ไม่ถูกต้อง");
                            } else {
                                gs.info(grDebtTask.number + " [Update] เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้ => [" + reason + "] " + reason_unable_help);
                                // grDebtTask.u_reason_unable_help = reason_unable_help;
                                // grDebtTask.update();
                            }
                        }
                    } else {
                        gs.info("[Error] ผู้ให้บริการไม่ตรงกัน Exce: "+rowExcel["ผู้ให้บริการ"] +" casetask: "+ grDebtTask.u_provider.getDisplayValue());
						
                    }
                } else {
                    gs.info("number diff!");
                    gs.info("[Error] ไม่สามารถ แก้ไข " + grDebtTask.number + "ด้วย script ได้");
					notMatchNumber.push(grDebtTask.sys_id.toString());
                }
                gs.info(grImportLog.target_record + ' แถวที่ ' + row + " แนวทางกับเหตุผล คือ " + rowExcel["แนวทางการช่วยเหลือลูกหนี้"] + " " + rowExcel["เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้"]);

                if(grDebtTask.next()){
					//do nothing
				}else{
					limitRow = 0;
					break;
				}
				allRecordSysId.push(grDebtTask.sys_id.toString());
                parser.next();

                rowCount++;
            }
            preV = parseInt(grImportLog.row);
        }
    }

    gs.info(text_log);
    text_log = "";

}

// log loop ของทั้งหมดของ case task ที่เราซ่อมไป 
gs.info(JSON.stringify(allRecordSysId));
gs.info("All record fix this script: "+allRecordSysId);
gs.info("All record can't fix this script: "+notMatchNumber);


function createExcelParser(detNumber) {
    var grImportLog = new GlideRecord("x_baot_debt_sett_0_import_excel_logs");
    // grImportLog.addEncodedQuery("target_record=" + detNumber + "^message=สำเร็จ^ORอัพเดตข้อมูลสำเร็จ");
    grImportLog.addEncodedQuery("target_record=" + detNumber + "^message=สำเร็จ");
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

                text_log += "\nRead Excel Attachment: " + grAtt.file_name;

                var attachment = new GlideSysAttachment();
                var attachmentStream = attachment.getContentStream(grAtt.sys_id);
                parser.setSource(attachmentStream);
                parser.setHeaderRowNumber(0);
                parser.parse();

            }

        }

		return false;
    } else{
		return true;
	}
}

function get_select_valid_active_check_fair(table_select, choice_col, choice_select, debtFair) {
    var grFair = new GlideRecord("u_project_debt");
    grFair.addEncodedQuery("u_name=" + debtFair);
    grFair.query();
    if (grFair.next()) {
        var selectGr = new GlideRecord(table_select);
        selectGr.addQuery(choice_col, choice_select);
        selectGr.addQuery("u_debt_fair", grFair.sys_id);
        selectGr.addQuery("u_active", "true");
        selectGr.setLimit(1);
        selectGr.query();
        if (selectGr.next()) {
            return selectGr.getUniqueValue();
        } else {
            return "Error";
        }
    }
}

function compareStrings(str1, str2, percent) {
    // ตรวจสอบความยาวของ string ทั้งสอง
    var len1 = str1.length;
    var len2 = str2.length;
 
    // สร้าง matrix เพื่อเก็บค่าความเหมือน
    var matrix = [];
 
    // เริ่มต้น matrix
    for (var i = 0; i <= len1; i++) {
        matrix[i] = [i];
    }
    for (var j = 0; j <= len2; j++) {
        matrix[0][j] = j;
    }
 
    // คำนวณความเหมือนด้วย Levenshtein Distance
    for (var i = 1; i <= len1; i++) {
        for (var j = 1; j <= len2; j++) {
            var cost = str1[i - 1] === str2[j - 1] ? 0 : 1;
            matrix[i][j] = Math.min(
                matrix[i - 1][j] + 1, // devarion
                matrix[i][j - 1] + 1, // insertion
                matrix[i - 1][j - 1] + cost // substitution
            );
        }
    }
 
    // คำนวณ percentage ของความเหมือน
    var maxLen = Math.max(len1, len2);
    var similarity = ((maxLen - matrix[len1][len2]) / maxLen) * 100;
 
    // ตรวจสอบว่า similarity ตรงไหม
    return similarity >= parseInt(percent);
}
 