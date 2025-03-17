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
grDebtTask.addEncodedQuery("");
grDebtTask.orderBy('number');
grDebtTask.setLimit(1000);
grDebtTask.query();
while (grDebtTask.next()) {

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
                gs.info(grImportLog.target_record+' เงินต้น แถวที่ ' + row + ' คือ: ' + MoneyValue+" เลขที่ยืนยันตัวตน คือ "+rowExcel["เลขที่ยืนยันตัวตน"]);
                //ใส่ค่าตรงนี้
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
 
