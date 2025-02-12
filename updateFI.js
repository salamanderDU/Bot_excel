var sheet_case_id = "";


if (sheet_case_id.substring(0, 3) == "DET") {

    var casetaskDT = new GlideRecord('x_baot_debt_sett_0_debt_task');
    casetaskDT.addQuery('number', sheet_case_id);
    casetaskDT.query()
    if (casetaskDT.next()) {

        casetaskDT.u_bulk_upload = true;
        target_record = sheet_case_task_id;


        if (isNotMember(userSysId_uploadby, casetaskDT.assignment_group)) {
            add_log(source.sys_import_set, row_count, target_record, "Error", "ไม่มีสิทธินำเข้าข้อมูลสำหรับผู้ให้บริการที่เลือก");
            error_log = error_log + 1;
        }
        else if (sheet_drd_report) {
            sheet_drd_report = formatDate(sheet_drd_report);
            if (check_date_valid(sheet_drd_report)) {
                casetaskDT.u_drd = set_date_object(sheet_drd_report);
            } else {
                add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด วันที่ทำสัญญา");
                error_log = error_log + 1;
            }
        }

        // Update Section
        if (error_log == 0) {
            casetaskDT.update();
            add_log(source.sys_import_set, row_count, target_record, "Update", "อัพเดทข้อมูลสำเร็จ วันที่ทำสัญญา");
        } else {
            error_log_all = 1;
            add_log(source.sys_import_set, row_count, target_record, "Error", "อัพเดตข้อมูลไม่สำเร็จ (Close Case) มีข้อผิดพลาด " + String(error_log) + " ข้อ");
        }
    }else {
        error_log_all = 1;
        add_log(source.sys_import_set, row_count, target_record, "Error", "อัพเดตข้อมูลไม่สำเร็จ หาเลขใบงาน");
    }

}





