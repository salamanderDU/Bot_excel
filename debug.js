gs.info(sheet_vehicle_number + ":" + sheet_vehicle_number_province);
if (carProduct.includes(sheet_product)) {
    if (sheet_vehicle_number) {
        if (sheet_vehicle_number_province) {
            casetaskGr.u_number_car = sheet_vehicle_number;
        } 
        // else {
        //     add_log(source.sys_import_set, row_count, target_record, "Error", "โปรดระบุข้อมูล จังหวัดที่จดทะเบียนรถ.");
        //     error_log = error_log + 1;
        // }
    } else if (sheet_vehicle_number_province) {
            add_log(source.sys_import_set, row_count, target_record, "Error", "โปรดระบุข้อมูล เลขที่ทะเบียนรถ");
            error_log = error_log + 1;
        }	// new
} else if (sheet_vehicle_number) {
        add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด ผลิตภัณฑ์ไม่ได้อยู่ในกลุ่มรถ");
        error_log = error_log + 1;
    }
//---------=-=-=-=-=-=-=-=-=-=-=-=-=-=--===-=--=-=--=-=--=-=-=--=-=-=-==-=-=-
if (carProduct.includes(sheet_product)) {
    if (sheet_vehicle_number_province) {
        if (sheet_vehicle_number) {
            var province_car = get_user_valid_active('u_car_registration_province', 'u_vehicle', sheet_vehicle_number_province);
            if (province_car == "Error") {
                add_log(source.sys_import_set, row_count, target_record, "Error", "รูปแบบข้อมูล จังหวัดที่จดทะเบียนรถ ไม่ถูกต้อง");
                error_log = error_log + 1;
            } else {
                casetaskGr.u_province_car = province_car;
            }
        } 
        // else {
        //     add_log(source.sys_import_set, row_count, target_record, "Error", "โปรดระบุข้อมูล เลขที่ทะเบียนรถ");
        //     error_log = error_log + 1;
        // }
    } else if (sheet_vehicle_number) {
            add_log(source.sys_import_set, row_count, target_record, "Error", "โปรดระบุข้อมูล จังหวัดที่จดทะเบียนรถ");
            error_log = error_log + 1;
        }	// new
} else if (sheet_vehicle_number_province) {
        add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด ผลิตภัณฑ์ไม่ได้อยู่ในกลุ่มรถ");
        error_log = error_log + 1;
    }