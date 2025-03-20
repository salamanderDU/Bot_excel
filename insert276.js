var text_log = "";
var lasrRecord;
var detNumber;
var row;
var cursor;
var importSetRef;
var parser = new sn_impex.GlideExcelParser();

var lastrow = 0;
var rowCount = 1;
var limitRow = 350;
// var recordExcel = 0;

var allRecordSysId = [];
var notMatchNumber = [];
var notMatchLog = [];
var errorLog = [];
var importSetrefcantFix = [];

var grDebtTask = new GlideRecord("x_baot_debt_sett_0_ect_debt_settlement");
grDebtTask.addEncodedQuery("sys_created_by=BulkUploadJob^u_glideline_debtISEMPTY^u_reason_unable_helpISEMPTY^u_resolution_codeIN100,200^state=1^assignment_group!=aeeedf8247f0d2101b2b4ed4116d4346^ORassignment_group=NULL");
grDebtTask.orderBy('sys_created_on');
// grDebtTask.orderBy('number');
grDebtTask.setLimit(limitRow);
grDebtTask.query();
while (rowCount < limitRow) {

    if (rowCount == 1) {
        grDebtTask.next();
        allRecordSysId.push(grDebtTask.sys_id.toString());
    }

    text_log += "\n" + grDebtTask.number + " " + grDebtTask.sys_created_on;
    detNumber = grDebtTask.number;
    cursor = 1;
    if (createExcelParser(detNumber)) {
        gs.info("[DEBUpdate]skip: " + grDebtTask.number + " " + grDebtTask.sys_id.toString());
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
                // gs.info("[DEBUpdate]cursor: " + cursor);
                // gs.info("logLastRow: " + logGetLastRow(importSetRef) + " excelRow: " + recordExcel);
                // if ((logGetLastRow(importSetRef) == recordExcel) && (grDebtTask.u_number_contact == rowExcel["เลขที่บัตร/ เลขที่สัญญา"]) && (grDebtTask.u_iden_number == rowExcel["เลขที่ยืนยันตัวตน"])) {
                if ((grDebtTask.u_number_contact == rowExcel["เลขที่บัตร/ เลขที่สัญญา"]) && (grDebtTask.u_iden_number == rowExcel["เลขที่ยืนยันตัวตน"])) {
					// gs.info("ผลการพิจารณา: "+rowExcel["ผลการพิจารณา"]);
                    if (rowExcel["ผลการพิจารณา"]) {
						var error_log = 0;
                        var simPercent = 40;
                        if (compareStrings(grDebtTask.u_provider.getDisplayValue(), rowExcel["ผู้ให้บริการ"]), simPercent) {


                            var sheet_result = rowExcel["ผลการพิจารณา"];
                            var debt_project = grDebtTask.u_debt_project.getDisplayValue();
                            var sheet_guidelines_debtors = rowExcel["แนวทางการช่วยเหลือลูกหนี้"];
							var sheet_reason_not_help = rowExcel["เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้"];
							

                            if (rowExcel["แนวทางการช่วยเหลือลูกหนี้"]) {
                                sheet_guidelines_debtors = sheet_guidelines_debtors.split("|")[0].trim();
                            } else {
                                sheet_guidelines_debtors = "";
                            }

                            if (rowExcel["เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้"]) {
                                sheet_reason_not_help = sheet_reason_not_help.split("|")[0].trim();
                            } else {
                                sheet_reason_not_help = "";
                            }

                            // แนวทางการช่วยเหลือลูกหนี้
                            if (sheet_result == "ได้ข้อสรุปกับลูกค้า") {
                                if (check_select_valid_activePipe2('u_glideline_debtor', "u_name=" + sheet_guidelines_debtors, debt_project)) {
                                    var glideline_debt = get_select_valid_active_check_fair('u_glideline_debtor', 'u_name', sheet_guidelines_debtors, debt_project);
                                    if (glideline_debt == "Error") {
                                        add_log(grDebtTask.number+ " "+grDebtTask.sys_id+" "+'รูปแบบข้อมูล "แนวทางการช่วยเหลือลูกหนี้" ไม่ถูกต้อง' + sheet_guidelines_debtors);
                                        error_log = error_log + 1;
                                    } else {
										
                                        gs.info(grDebtTask.number + " [Update] แนวทางการช่วยเหลือลูกหนี้ => [" + glideline_debt + "] ");
                                        grDebtTask.u_glideline_debt = glideline_debt;
                                        // grDebtTask.update();
                                    }
                                } else if (sheet_guidelines_debtors == "") {
                                    add_log(grDebtTask.number+ " "+grDebtTask.sys_id+" "+'โปรดระบุ "แนวทางการช่วยเหลือลูกหนี้"' + sheet_guidelines_debtors);
                                    error_log = error_log + 1;
                                } else {
                                    add_log(grDebtTask.number+ " "+grDebtTask.sys_id+" "+'โปรดระบุ "แนวทางการช่วยเหลือลูกหนี้" ให้ถูกต้องตามโครงการ' + sheet_guidelines_debtors);
                                    error_log = error_log + 1;
                                }
                            }


                            // เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้
                            if (sheet_result == "ไม่ได้ข้อสรุปกับลูกค้า") {
                                if (check_select_valid_activePipe2('u_unable_help', "u_name=" + sheet_reason_not_help, debt_project)) {
                                    var reason_unable_help = get_select_valid_active_check_fair('u_unable_help', 'u_name', sheet_reason_not_help, debt_project);
                                    if (reason_unable_help == "Error") {
                                        add_log(grDebtTask.number+ " "+grDebtTask.sys_id+" "+'รูปแบบข้อมูล "เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้" ไม่ถูกต้อง');
                                        error_log = error_log + 1;
                                    } else {
                                        gs.info(grDebtTask.number + " [Update] เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้ => [" + reason_unable_help + "] " );
                                        grDebtTask.u_reason_unable_help = reason_unable_help;
                                    }
                                } else if (sheet_reason_not_help == "") {
                                    add_log(grDebtTask.number+ " "+grDebtTask.sys_id+" "+'โปรดระบุ "เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้"');
                                    error_log = error_log + 1;
                                } else {
                                    add_log(grDebtTask.number+ " "+grDebtTask.sys_id+" "+'โปรดระบุ "เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้" ให้ถูกต้องตามโครงการ');
                                    error_log = error_log + 1;

                                }
                            }

							//ข้อมูลที่ต้องการแจ้ง ธปท. เพิ่มเติม (ถ้ามี)
							if((grDebtTask.close_notes == "") && (rowExcel["ข้อมูลที่ต้องการแจ้ง ธปท. เพิ่มเติม (ถ้ามี)"] != "")){
								grDebtTask.close_notes = rowExcel["ข้อมูลที่ต้องการแจ้ง ธปท. เพิ่มเติม (ถ้ามี)"];
                                gs.info("ข้อมูลที่ต้องการแจ้ง ธปท. เพิ่มเติม (ถ้ามี) => [" + rowExcel["ข้อมูลที่ต้องการแจ้ง ธปท. เพิ่มเติม (ถ้ามี)"] + "]");
							}


                        } else {
                            add_log(grDebtTask.number+ " "+grDebtTask.sys_id+" "+"ผู้ให้บริการไม่ตรงกัน Exce: " + rowExcel["ผู้ให้บริการ"] + " casetask: " + grDebtTask.u_provider.getDisplayValue());

                        }
						
						if(error_log == 0){
							gs.info(grDebtTask.number+ " "+"update");
						}else{
							add_log(grDebtTask.number+ " "+grDebtTask.sys_id+" "+"มีข้อผิดพลาด "+error_log+" ข้อ");
						}

                    } else {
                        add_log(grDebtTask.number+ " "+grDebtTask.sys_id+" "+"โปรดระบุ ผลการพิจารณา");
                        notMatchNumber.push(grDebtTask.sys_id.toString());
                    }
                } else {
                    gs.info("ข้อมูลไม่ตรง");
					importSetrefcantFix.push(importSetRef.toString());
					notMatchLog.push(grDebtTask.sys_id.toString());
                }
                gs.info(grImportLog.target_record + ' แถวที่ ' + row + " แนวทางกับเหตุผล คือ " + rowExcel["แนวทางการช่วยเหลือลูกหนี้"] + " " + rowExcel["เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้"]);
                // gs.info(grImportLog.target_record + ' แถวที่ ' + row + " เลขที่ยืนยันตัวตน คือ " + rowExcel["เลขที่ยืนยันตัวตน"]);
                // gs.info(grImportLog.target_record + ' แถวที่ ' + row + " เลขที่บัตร/ เลขที่สัญญา คือ " + rowExcel["เลขที่บัตร/ เลขที่สัญญา"]);


                if (grDebtTask.next()) {
                    //do nothing
                } else {
                    parser.close();
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
gs.info("[DEBUpdate]All record fix this script: " + allRecordSysId);
importSetrefcantFix = unique(importSetrefcantFix);
gs.info("[DEBUpdate]All Import set log wrong: " + unique(importSetrefcantFix));
gs.info("[DEBUpdate]All record fix this script: " + notMatchLog);
gs.info("[DEBUpdate]All record error log: " + errorLog);



function createExcelParser(detNumber) {
    var grImportLog = new GlideRecord("x_baot_debt_sett_0_import_excel_logs");
    grImportLog.addEncodedQuery("target_record=" + detNumber + "^stateSTARTSWITHinsert");
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

                // var attachmentRowCount = new GlideSysAttachment();
                // var attachmentStreamRowCount = attachment.getContentStream(grAtt.sys_id);
                // var parserRowCount = new sn_impex.GlideExcelParser();
                // parserRowCount.setSource(attachmentStreamRowCount);
                // parserRowCount.setHeaderRowNumber(0);
                // parserRowCount.parse();
                // var count = 1;
                // while (parserRowCount.next()) {
                //     count += 1;
                // }
                // parserRowCount.close();

                // recordExcel = count;
                // gs.info("[All excel count] = " + recordExcel);


            }

        }

        return false;
    } else {
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

function logGetLastRow(importsetref) {
    var grImportLog = new GlideRecord("x_baot_debt_sett_0_import_excel_logs");
    grImportLog.addEncodedQuery("importsetref=" + importSetRef);
    grImportLog.orderByDesc('row');
    grImportLog.query();
    if (grImportLog.next()) {
        return grImportLog.row;
    }
}

function check_select_valid_activePipe2(table_select, enCodeQuery, debtFair) {
	var grFair = new GlideRecord("u_project_debt");
	grFair.addEncodedQuery("u_name=" + debtFair);
	grFair.query();
	if (grFair.next()) {
		//query agian
		var selectGr = new GlideRecord(table_select);
		selectGr.addEncodedQuery(enCodeQuery);
		selectGr.addQuery("u_debt_fair", grFair.sys_id);
		selectGr.setLimit(1);
		selectGr.query();
		if (selectGr.next()) {
			return true;
		} else {
			return false;
		}
	}
}

function add_log(msg){
	gs.info(msg);
	errorLog.push(msg);
}

function get_active_result(choice_select) {
    var selectGr = new GlideRecord('sys_choice');
    selectGr.addQuery('name', 'u_debt_settlement_master_case_task');
    selectGr.addQuery('label', choice_select);
    selectGr.addQuery('element', 'u_resolution_code');
    selectGr.setLimit(1);
    selectGr.query();
    if (selectGr.next()) {
        return selectGr.value;

    } else {
        return "Error";
    }
}

function unique(a){
    var a1 = a.filter((e, i, self) => i === self.indexOf(e));
    return a1
}