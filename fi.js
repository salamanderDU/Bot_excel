(function runTransformScript(source, map, log, target) {

    var importExcelAttachmentSysId = "";
    var userSysId_uploadby = "";
    try {

        var targetTable = source.getTableName();
        var targetGR = new GlideRecord(targetTable);
        targetGR.addQuery('sys_import_set', source.sys_import_set);
        targetGR.query();
        var count = targetGR.getRowCount();
        var row_count = 1;
        var error_log = 0;
        var error_log_all = 0;
        var target_record = "";

        var carProduct = ["สินเชื่อเช่าซื้อรถยนต์ / Top-up / รถแลกเงินแบบโอนเล่ม", "สินเชื่อเช่าซื้อรถจักรยานยนต์ / Top-up / รถแลกเงินโอนเล่ม", "จำนำทะเบียนรถยนต์ (ไม่โอนเล่ม)", "จำนำทะเบียนรถจักรยานยนต์ (ไม่โอนเล่ม)", "สินเชื่อเช่าซื้อรถยนต์", "สินเชื่อเช่าซื้อรถจักรยานยนต์", "สินเชื่อจำนำทะเบียนรถยนต์", "สินเชื่อจำนำทะเบียนรถจักรยานยนต์"];


        //upload by
        // เรียก upload by ผ่าน attachment บน data source ที่มีค่าของ record IMP
        var grSA = new GlideRecord('sys_attachment');
        grSA.addEncodedQuery("table_sys_id=" + source.sys_import_set.data_source.sys_id);
        grSA.setLimit(1);
        grSA.orderByDesc("sys_created_on");
        grSA.query();
        if (grSA.next()) {
            var sysid_table_importattachment1 = grSA.getValue("file_name").split("-")[0];
            var row_table_importattachment1 = new GlideRecord('x_baot_debt_sett_0_import_excel_attachments');
            row_table_importattachment1.addQuery('sys_id', sysid_table_importattachment1);
            row_table_importattachment1.query();
            if (row_table_importattachment1.next()) {
                userSysId_uploadby = row_table_importattachment1.upload_by;
                importExcelAttachmentSysId = row_table_importattachment1.sys_id;
            }
        }

        while (targetGR.next()) {
            row_count += 1;
            error_log = 0;
            target_record = "-";


            var codeDiff = [0,  1,  3,  4,  9, 18, 22, 24, 25, 26, 27];

            // ขอรับบริการในนาม 0
            // var sheet_request_service_on_behalf_of = targetGR.u_ขอร_บบร_การในนาม;
            // ผู้ติดต่อ 1
            // var sheet_contact = targetGR.u_ผ__ต_ดต_อ;
            // ผู้ขอรับบริการ 2
            var sheet_service_requester = targetGR.u_ผ__ขอร_บบร_การ.toString();
            // ประเภทเลขที่ยืนยันตัวตน 3
            // var sheet_type_identification_number = targetGR.u_ประเภทเลขท__ย_นย_นต_วตน;
            // ประเภทเลขที่ยืนยันตัวตน (นิติบุคคล) 4
            // var sheet_type_identification_number_juristic = targetGR.u_ประเภทเลขท___ตน__น_ต_บ_คคล_;
            // เลขที่ยืนยันตัวตน 5
            var sheet_identification_number = targetGR.u_เลขท__ย_นย_นต_วตน.toString();
            // เลขที่ยืนยันตัวตน (นิติบุคคล) 6
            var sheet_identification_number_juristic = targetGR.u_เลขท__ย_นย_น_ตน__น_ต_บ_คคล_.toString();
            // หมายเลขโทรศัพท์ 7
            var sheet_phone = targetGR.u_หมายเลขโทรศ_พท_.toString();
            // อีเมล 8
            var sheet_email = targetGR.u_อ_เมล.toString();
            // จังหวัด (ที่อยู่ลูกหนี้) 9
            // var sheet_province = targetGR.u_จ_งหว_ด__ท__อย__ล_กหน___;
            var codeDiff = [0,  1,  3,  4,  9, 18, 22, 24, 25, 26, 27];
            // ชื่อนิติบุคคล 10
            var sheet_juristic_name = targetGR.u_ช__อน_ต_บ_คคล.toString();
            // ผู้ให้บริการ 11
            var sheet_provider = targetGR.u_ผ__ให_บร_การ;
            // ผลิตภัณฑ์ 12
            var sheet_product = targetGR.u_ผล_ตภ_ณฑ_;
            // สถานะบัญชี 13
            var sheet_account_status = targetGR.u_สถานะบ_ญช_;
            // สถานะบัญชี 13.5
            var sheet_project_debt = "โครงการคุณสู้ เราช่วย";
            // เลขที่ทะเบียนรถ 14
            var sheet_vehicle_number = targetGR.u_เลขท__ทะเบ_ยนรถ;
            // จังหวัดที่จดทะเบียนรถ 15
            var sheet_vehicle_number_province = targetGR.u_จ_งหว_ดท__จดทะเบ_ยนรถ;
            // แนวทางที่ต้องการเสนอให้ผู้บริการทางการเงินพิจารณา 16
            var sheet_guidelines_provider = targetGR.u_แนวทางท__ต_อ_การเง_นพ_จารณา;
            // เลขที่บัตร/ เลขที่สัญญา 17
            var sheet_contract_number = targetGR.u_เลขท__บ_ตร__เลขท__ส_ญญา.toString();
            // วันที่เริ่มติดต่อลูกหนี้  18
            // var sheet_date_of_start = targetGR.u_ว_นท__เร__มต_ดต_อล_กหน__;
            // ผลการพิจารณา 19
            var sheet_result = targetGR.u_ผลการพ_จารณา.toString();
            // แนวทางการช่วยเหลือลูกหนี้ 20
            var sheet_guidelines_debtors = targetGR.u_แนวทางการช_วยเหล_อล_กหน__;
            // เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้ 21
            var sheet_reason_not_help = targetGR.u_เหต_ผลท__ไม__หล_อล_กหน__ได_;
            // ภาระหนี้รวม (บาท) 22
            // var sheet_total_debt = targetGR.u_ภาระหน__รวม__บาท_;
            var codeDiff = [0,  1,  3,  4,  9, 18, 22, 24, 25, 26, 27];
            // เงินต้น (บาท) 23
            var sheet_principal = targetGR.u_เง_นต_น__บาท_;
            // ภาระหนี้ที่ตกลงชำระ (บาท) 24
            // var sheet_debt_agreed = targetGR.u_ภาระหน__ท__ตกลงชำระ__บาท_;
            // จำนวนงวดที่ชำระ (เดือน) 25
            // var sheet_no_payment = targetGR.u_จำนวนงวดท__ชำระ__เด_อน_.toString();
            // ค่างวดต่อเดือน (บาท) 26
            // var sheet_monthly = targetGR.u_ค_างวดต_อเด_อน__บาท_;
            // รายงาน RDT 27
            // var sheet_rtd_report = targetGR.u_รายงาน_rdt;
            // วันที่ทำสัญญา 28
            var sheet_contract_date = targetGR.u_ว_นท__ทำส_ญญา;
            // ข้อมูลที่ต้องการแจ้ง ธปท. เพิ่มเติม (ถ้ามี) 29
            var sheet_additional_info = targetGR.u_ข_อม_ลท__ต_อ__มเต_ม__ถ_าม__;
            //วันที่รับคำขอ 30
            var sheet_take_request = targetGR.u_ว_นท__ร_บคำขอ;

            //ตัดค่า | ที่ระบุถึงโครงการแก้หนี้
            if (sheet_product) {
                sheet_product = targetGR.u_ผล_ตภ_ณฑ_.split("|")[0].trim();
            }
            if (sheet_account_status) {
                sheet_account_status = targetGR.u_สถานะบ_ญช_.split("|")[0].trim();
            }
            if (sheet_guidelines_debtors) {
                sheet_guidelines_debtors = targetGR.u_แนวทางการช_วยเหล_อล_กหน__.split("|")[0].trim();
            }
            if (sheet_reason_not_help) {
                sheet_reason_not_help = targetGR.u_เหต_ผลท__ไม__หล_อล_กหน__ได_.split("|")[0].trim();
            }
            // ตรววจสอบ format ข้อมูลวันที่
            // sheet_rtd_report = formatDate(sheet_rtd_report);


            var casetaskGr = new GlideRecord('x_baot_debt_sett_0_debt_task');
            casetaskGr.initialize();
            casetaskGr.u_bulk_upload = true;
            // warning
            casetaskGr.u_show_on_ticket_list_page = true;

            // ขอรับบริการในนาม
            // if (sheet_request_service_on_behalf_of) {
            //     casetaskGr.u_walkin_receive_service = sheet_request_service_on_behalf_of;
            // } else {
            //     add_log(source.sys_import_set, row_count, target_record, "Error", "โปรดระบุ ขอรับบริการในนาม");
            //     error_log = error_log + 1;
            // }
            // ผู้ติดต่อ
            // if (sheet_contact) {
            //     casetaskGr.u_walkin_contact = sheet_contact; //new req
            // } else {
            //     add_log(source.sys_import_set, row_count, target_record, "Error", "โปรดระบุ ผู้ติดต่อ");
            //     error_log = error_log + 1;
            // }
            // ผู้ขอรับบริการ
            // validate ว่าเป็นชื่อหรือไม่ ได้ thai และ อังกฤษ และชื่อไม่เป็น number หรือค่าว่าง
            if (validateName(sheet_service_requester)) {
                casetaskGr.u_walk_service_requester = sheet_service_requester;
            } else {
                add_log(source.sys_import_set, row_count, target_record, "Error", "รูปแบบข้อมูล ผู้ขอรับบริการ ไม่ถูกต้อง");
                error_log = error_log + 1;
            }
            // ประเภทเลขที่ยืนยันตัวตน
            //sheet_type_identification_number ระบบให้เป็น เลขบัตรแล้ว 
            // if (sheet_type_identification_number) {
            //     casetaskGr.u_walkin_type_identification_number = sheet_type_identification_number;
            // }

            // ประเภทเลขที่ยืนยันตัวตน (นิติบุคคล)
            // if (sheet_type_identification_number_juristic) {
            //     var typeIdentificationNumberJuristic = getRef("u_type_verify_iden", "u_verify_identify", sheet_type_identification_number_juristic);
            //     if (typeIdentificationNumberJuristic == "Error") {
            //         add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด ประเภทเลขที่ยืนยันตัวตน (นิติบุคคล)");
            //         error_log = error_log + 1;
            //     } else {
            //         casetaskGr.u_walkin_type_identification_number_juristic = typeIdentificationNumberJuristic;
            //     }
            // } else {
            //     add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด ประเภทเลขที่ยืนยันตัวตน (นิติบุคคล) ไม่มีข้อมูล");
            //     error_log = error_log + 1;
            // }
            // เลขที่ยืนยันตัวตน
            // ไม่เป็น string หรือ ค่าว่าง | Error, รูปแบบข้อมูล "เลขที่ยืนย้นตัวตน" ไม่ถูกต้อง
            if (validateNumber(sheet_identification_number)) {
                casetaskGr.u_walkin_identification_number = sheet_identification_number;
            } else {
                add_log(source.sys_import_set, row_count, target_record, "Error", "รูปแบบข้อมูล เลขที่ยืนยันตัวตน ไม่ถูกต้อง");
                error_log = error_log + 1;
            }
            // เลขที่ยืนยันตัวตน (นิติบุคคล)
            // มีค่าแต่ไม่เป็น String	Error, รูปแบบข้อมูล "เลขที่ยืนยันตัวตน (นิติบุคคล)" ไม่ถูกต้อง
            // เป็น Empty แต่"ชื่อนิติบุคคล" not empty	Error, โปรดระบุ "เลขที่ยืนยันตัวตน (นิติบุคคล)" 
            if (sheet_identification_number_juristic == "" && sheet_juristic_name != "") {
                add_log(source.sys_import_set, row_count, target_record, "Error", "โปรดระบุ เลขที่ยืนยันตัวตน (นิติบุคคล)");
                error_log = error_log + 1;
            } else if (validateNumber(sheet_identification_number_juristic)) {
                casetaskGr.u_identification_number_juristic = sheet_identification_number_juristic;
            } else {
                add_log(source.sys_import_set, row_count, target_record, "Error", "รูปแบบข้อมูล เลขที่ยืนยันตัวตน (นิติบุคคล) ไม่ถูกต้อง");
                error_log = error_log + 1;
            }

            // หมายเลขโทรศัพท์
            // เป็น Empty	Error, โปรดระบุข้อมูล "หมายเลขโทรศัพท์" ในรูปแบบ E164 เช่น +66812345678
            // ไม่เป็น E164   Error, รูปแบบข้อมูล "หมายเลขโทรศัพท์" ไม่เป็น E164 เช่น +66812345678
            if (sheet_phone == "") {
                add_log(source.sys_import_set, row_count, target_record, "Error", "โปรดระบุข้อมูล หมายเลขโทรศัพท์ ในรูปแบบ E164 เช่น +66812345678");
                error_log = error_log + 1;
            } else if (validatePhone(sheet_phone)) {
                casetaskGr.u_walkin_phone = sheet_phone;
            } else {
                add_log(source.sys_import_set, row_count, target_record, "Error", "รูปแบบข้อมูล หมายเลขโทรศัพท์ ไม่เป็น E164 เช่น +66812345678");
                error_log = error_log + 1;
            }

            // อีเมล
            // เป็น Empty	Error, โปรดระบุข้อมูล "อีเมล" 
            // ไม่เป็น Email format   Error, รูปแบบข้อมูล "อีเมล" ไม่เป็นตามรูปแบบ เช่น example@domain.xyz
            // มีค่าตรงตาม format แต่ไม่มีอยู่ในระบบหรือในกลุ่มตาม Assigned Group ของ Assignment Record	Error, ไม่พบข้อมูลผู้ใช้ Email บนระบบ
            if (sheet_email == "") {
                add_log(source.sys_import_set, row_count, target_record, "Error", "โปรดระบุข้อมูล อีเมล");
                error_log = error_log + 1;
            } else if (validateEmail(sheet_email)) {
                casetaskGr.u_walkin_email = sheet_email;
            } else {
                add_log(source.sys_import_set, row_count, target_record, "Error", "รูปแบบข้อมูล อีเมล ไม่เป็นตามรูปแบบ เช่น example@domain.xyz");
                error_log = error_log + 1;
            }

            // จังหวัด (ที่อยู่ลูกหนี้)
            // if (sheet_province) {
            //     var province = getRef("u_province", "u_name", sheet_province);
            //     if (province == "Error") {
            //         add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด จังหวัด (ที่อยู่ลูกหนี้)");
            //         error_log = error_log + 1;
            //     } else {
            //         casetaskGr.u_province = province;
            //     }
            // } else {
            //     add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด จังหวัด (ที่อยู่ลูกหนี้)");
            //     error_log = error_log + 1;
            // }
            // ชื่อนิติบุคคล
            // มีค่าแต่ไม่เป็น String	Error, รูปแบบข้อมูล "ชื่อนิติบุคคล" ไม่ถูกต้อง
            // เป็น Empty แต่"เลขที่ยืนยันตัวตน (นิติบุคคล)" not empty	Error, โปรดระบุ "ชื่อนิติบุคคล" 

            if (sheet_juristic_name == "" && sheet_identification_number_juristic != "") {
                add_log(source.sys_import_set, row_count, target_record, "Error", "โปรดระบุ ชื่อนิติบุคคล");
                error_log = error_log + 1;
            } else if (validateString(sheet_juristic_name)) {
                casetaskGr.u_walkin_juristic_name = sheet_juristic_name;
            } else {
                add_log(source.sys_import_set, row_count, target_record, "Error", "รูปแบบข้อมูล ชื่อนิติบุคคล ไม่ถูกต้อง");
                error_log = error_log + 1;
            }

            // ผู้ให้บริการ
            // เป็น Empty	Error, โปรดระบุ "ผู้ให้บริการ" 
            // มีค่าแต่ไม่ตรงกับที่มีบนระบบ	Error, รูปแบบข้อมูล "ผู้ให้บริการ"  ไม่ถูกต้อง
            // มีค่าบนระบบแต่ผู้อัพโหลดไม่มีสิทธิในการอัพโหลด	Error, ไม่มีสิทธินำเข้าข้อมูลสำหรับผู้ให้บริการที่เลือก
            if (sheet_provider == "") {
                add_log(source.sys_import_set, row_count, target_record, "Error", "โปรดระบุ ผู้ให้บริการ");
                error_log = error_log + 1;
            } else {
                var providerSupportGroup;
                var grCTA = new GlideRecord("x_baot_complaint_case_task_assignment");
                grCTA.addEncodedQuery("u_provider.u_display_name=" + sheet_provider);
                grCTA.query();
                if (grCTA.next()) {

                    providerSupportGroup = grCTA.u_support_group.sys_id;
                    var dev = false;
                    if (isNotMember(userSysId_uploadby, providerSupportGroup)) {
                        add_log(source.sys_import_set, row_count, target_record, "Error", "ไม่มีสิทธินำเข้าข้อมูลสำหรับผู้ให้บริการที่เลือก");
                        error_log = error_log + 1;
                    } else {
                        var provider = getRef("u_m2m_provider_to_application", "u_provider.u_name", sheet_provider);
                        if (provider == "Error") {
                            add_log(source.sys_import_set, row_count, target_record, "Error", "รูปแบบข้อมูล ผู้ให้บริการ ไม่ถูกต้อง");
                            error_log = error_log + 1;
                        } else {
                            casetaskGr.u_provider = provider;
                        }
                    }
                } else {
                    add_log(source.sys_import_set, row_count, target_record, "Error", "รูปแบบข้อมูล ผู้ให้บริการ ไม่ถูกต้อง");
                    error_log = error_log + 1;
                }
            }

            // ผลิตภัณฑ์
            // เป็น Empty	Error, โปรดระบุ "ผลิตภัณฑ์" 
            // มีค่าแต่ไม่ตรงกับที่มีบนระบบ	Error, รูปแบบข้อมูล "ผลิตภัณฑ์"  ไม่ถูกต้อง
            // มีค่าบนระบบแต่ไม่ถูกต้องตามโครงการ	Error, โปรดระบุ "ผลิตภัณฑ์" ให้ถูกต้องตามโครงการ
            if (sheet_product == "") {
                add_log(source.sys_import_set, row_count, target_record, "Error", "โปรดระบุ ผลิตภัณฑ์");
                error_log = error_log + 1;
            } else if (check_select_valid_activePipe('x_baot_debt_sett_0_product_to_fair', "u_m2m_product_to_application.u_productSTARTSWITH" + sheet_product + "^x_baot_debt_sett_0_debt_fair.short_descriptionLIKE" + sheet_project_debt)) {
                var product = get_select_valid_active('u_m2m_product_to_application', 'u_product.u_product', sheet_product);
                if (product == "Error") {
                    add_log(source.sys_import_set, row_count, target_record, "Error", "รูปแบบข้อมูล ผลิตภัณฑ์ ไม่ถูกต้อง");
                    error_log = error_log + 1;
                } else {
                    casetaskGr.u_product = product;
                }

            } else {
                add_log(source.sys_import_set, row_count, target_record, "Error", "โปรดระบุ ผลิตภัณฑ์ ให้ถูกต้องตามโครงการ");
                error_log = error_log + 1;

            }

            // สถานะบัญชี
            // เป็น Empty	Error, โปรดระบุ "สถานะบัญชี" 
            // มีค่าแต่ไม่ตรงกับที่มีบนระบบ	Error, รูปแบบข้อมูล "สถานะบัญชี"  ไม่ถูกต้อง
            // มีค่าบนระบบแต่ไม่ถูกต้องตามโครงการ	Error, โปรดระบุ "สถานะบัญชี" ให้ถูกต้องตามโครงการ
            if (sheet_account_status == "") {
                add_log(source.sys_import_set, row_count, target_record, "Error", "โปรดระบุ สถานะบัญชี");
                error_log = error_log + 1;
            } else if (sheet_account_status == "" && sheet_state == "Cancelled") {
                // do nothing
            } else if (check_select_valid_activePipe('x_baot_debt_sett_0_debt_to_fair', "u_m2m_debt_status_to_application.u_debt_statusSTARTSWITH" + sheet_account_status + "^x_baot_debt_sett_0_debt_fair.short_description=" + sheet_project_debt)) {
                var state_debt = get_select_valid_active('u_m2m_debt_status_to_application', 'u_debt_status.u_name', sheet_account_status);
                if (state_debt == "Error") {
                    add_log(source.sys_import_set, row_count, target_record, "Error", "รูปแบบข้อมูล สถานะบัญชี ไม่ถูกต้อง");
                    error_log = error_log + 1;
                } else {
                    casetaskGr.u_state_debt = state_debt;
                }
            } else {
                add_log(source.sys_import_set, row_count, target_record, "Error", "โปรดระบุ สถานะบัญชี ให้ถูกต้องตามโครงการ");
                error_log = error_log + 1;
            }
            // โครงการแก้หนี้
            // อันนี้เป็นของโครงการคุณสู้เราช่วย
            if (sheet_project_debt) {
                casetaskGr.u_debt_project = "d343d5cb1bedc6506c24dcace54bcb85"; // โครงการคุณสู้ เราช่วย
            }
            // เลขที่ทะเบียนรถ
            // เป็น Empty แต่ จังหวัดที่จดทะเบียนรถ not empty	Error, โปรดระบุข้อมูล "เลขที่ทะเบียนรถ" 
            if (carProduct.includes(sheet_product)) {
                if (sheet_vehicle_number == "" && sheet_vehicle_number_province != "") {
                    add_log(source.sys_import_set, row_count, target_record, "Error", "โปรดระบุข้อมูล เลขที่ทะเบียนรถ");
                    error_log = error_log + 1;
                } else {
                    // ห้ามเป้นค่าว่าง ถ้าเป็นสินเชื่อเช่าซื้อรถยนต์, สินเชื่อเช่าซื้อรถจักรยานยนต์, สินเชื่อจำนำทะเบียนรถยนต์, สินเชื่อจำนำทะเบียนรถจักรยานยนต์
                    if (sheet_vehicle_number != "") {
                        casetaskGr.u_number_car = sheet_vehicle_number;
                    } else {
                        add_log(source.sys_import_set, row_count, target_record, "Error", "โปรดระบุข้อมูล เลขที่ทะเบียนรถ");
                        error_log = error_log + 1;
                    }
                }
            }
            // จังหวัดที่จดทะเบียนรถ
            // เเป็น Empty แต่ เลขที่ทะเบียนรถ not empty	Error, โปรดระบุข้อมูล "จังหวัดที่จดทะเบียนรถ" 
            // มีค่าแต่ไม่ตรงกับค่าในระบบ	Error, รูปแบบข้อมูล "จังหวัดที่จดทะเบียนรถ"ไม่ถูกต้อง 
            if (carProduct.includes(sheet_product)) {
                if (sheet_vehicle_number != "") {
                    var province_car = get_user_valid_active('u_car_registration_province', 'u_vehicle', sheet_vehicle_number_province);
                    if (province_car == "Error") {
                        add_log(source.sys_import_set, row_count, target_record, "Error", "รูปแบบข้อมูล จังหวัดที่จดทะเบียนรถ ไม่ถูกต้อง");
                        error_log = error_log + 1;
                    } else {
                        casetaskGr.u_province_car = province_car;
                    }
                } else {
                    add_log(source.sys_import_set, row_count, target_record, "Error", "โปรดระบุข้อมูล จังหวัดที่จดทะเบียนรถ");
                    error_log = error_log + 1;
                }
            }

            // แนวทางที่ต้องการเสนอให้ผู้บริการทางการเงินพิจารณา
            // ถ้ามีค่า แต่ไม่ตรงกับค่าที่ควรจะเป็น	Error, รูปแบบข้อมูล "แนวทางที่ต้องการเสนอให้ผู้บริการทางการเงินพิจารณา"ไม่ถูกต้อง 
            if (sheet_guidelines_provider) {
                var glideline_provider = getRef("u_glideline_debts", "u_name", sheet_guidelines_provider);
                if (glideline_provider == "Error") {
                    add_log(source.sys_import_set, row_count, target_record, "Error", "รูปแบบข้อมูล แนวทางที่ต้องการเสนอให้ผู้บริการทางการเงินพิจารณา");
                    error_log = error_log + 1;
                } else {
                    casetaskGr.u_offer_provider = glideline_provider; //old req ref
                }
            }
            // เลขที่บัตร/ เลขที่สัญญา
            // ไม่เป็น String หรือ Empty	Error, รูปแบบข้อมูล "เลขที่บัตร/ เลขที่สัญญา" ไม่ถูกต้อง
            if (validateString(sheet_contract_number)) {
                casetaskGr.u_number_contact = sheet_contract_number;
            } else {
                add_log(source.sys_import_set, row_count, target_record, "Error", "รูปแบบข้อมูล เลขที่บัตร/ เลขที่สัญญา ไม่ถูกต้อง");
                error_log = error_log + 1;
            }
            // วันที่เริ่มติดต่อลูกหนี้
            // if (check_datetime_valid(sheet_date_of_start)) {
            //     casetaskGr.u_contact_debt = set_timezone_object(sheet_date_of_start);
            // } else if (sheet_date_of_start == "") {
            //     add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด วันที่เริ่มติดต่อลูกหนี้ไม่มีข้อมูล");
            //     error_log = error_log + 1;
            // } else {
            //     add_log(source.sys_import_set, row_count, target_record, "Error", "รูปแบบข้อมูล วันที่เริ่มติดต่อลูกหนี้ ไม่ถูกต้อง");
            //     error_log = error_log + 1;
            // }

            // ผลการพิจารณา ติดไว้ก่อน รอมัดรวมกับ state 
            //เป็น Empty	Error, โปรดระบุ "ผลการพิจารณา" 
            // มีค่าแต่ไม่ตรงกับที่มีบนระบบ	Error, รูปแบบข้อมูล "ผลการพิจารณา"  ไม่ถูกต้อง
            
            var cancelledResolution = ["ไม่พบบัญชีลูกค้า", "บัญชีถูกขายไปแล้ว", "ยอดหนี้เท่ากับศูนย์", "ปิดบัญชีแล้ว", "คำขอซ้ำ", "ลูกค้ายกเลิกคำขอ", "ยกเลิกโดยเจ้าหน้าที่"];
            var resolvedResolution = ["ได้ข้อสรุปกับลูกค้า", "ไม่ได้ข้อสรุปกับลูกค้า"];

            // ได้ข้อสรุปกับลูกค้า
            // ไม่ได้ข้อสรุปกับลูกค้า
            // เป็นค่าว่าง    Error, โปรดระบุ ผลการพิจารณา
            // มีค่าแต่ไม่ตรงกับที่มีบนระบบ	Error, รูปแบบข้อมูล "ผลการพิจารณา"  ไม่ถูกต้อง
            // ไม่มีข้อมูลเพียงพอสำหรับการนำข้อมูลผลการพิจารณา	Error, ผิดพลาด ข้อมูลไม่เพียงพอสำหรับการนำข้อมูลผลการพิจารณา: ได้ข้อสรุปกับลูกค้า
            if (sheet_result == "ได้ข้อสรุปกับลูกค้า") {
                if (sheet_guidelines_debtors == "") {
                    add_log(source.sys_import_set, row_count, target_record, "Error", "โปรดระบุ แนวทางการช่วยเหลือลูกหนี้");
                    error_log = error_log + 1;
                } else {
                    var glideline_debt = get_select_valid_active_check_fair('u_glideline_debtor', 'u_name', sheet_guidelines_debtors, sheet_project_debt);
                    if (glideline_debt == "Error" || (sheet_product == '' || sheet_account_status == '' || sheet_contract_number == '' || sheet_principal == '')) {
                        add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด ข้อมูลไม่เพียงพอสำหรับการนำข้อมูลผลการพิจารณา: " + sheet_result);
                        error_log = error_log + 1;
                    } else {
                        casetaskGr.u_glideline_debt = glideline_debt;
                        // resolved
                        casetaskGr.u_resolution_code = 100;
                        casetaskGr.state = 6; // Resolved
                        casetaskGr.u_work_state = "เจ้าหน้าที่ดำเนินการเสร็จสิ้น";
                        casetaskGr.close_notes = sheet_additional_info;
                    }
                }
            }
            else if (sheet_result == "ไม่ได้ข้อสรุปกับลูกค้า") {
                if (sheet_reason_not_help == "") {
                    add_log(source.sys_import_set, row_count, target_record, "Error", "โปรดระบุ เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้");
                    error_log = error_log + 1;
                } else {
                    var reason_unable_help = get_select_valid_active_check_fair('u_unable_help', 'u_name', sheet_reason_not_help, sheet_project_debt);
                    if (reason_unable_help == "Error" || (sheet_product == '' || sheet_account_status == '' || sheet_contract_number == '' || sheet_principal == '')) {
                        add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด ข้อมูลไม่เพียงพอสำหรับการนำข้อมูลผลการพิจารณา: " + sheet_result);
                        error_log = error_log + 1;
                    } else {
                        casetaskGr.u_reason_unable_help = reason_unable_help;
                        // resolved
                        casetaskGr.u_resolution_code = 200;
                        casetaskGr.state = 6; // Resolved
                        casetaskGr.u_work_state = "เจ้าหน้าที่ดำเนินการเสร็จสิ้น";
                        casetaskGr.close_notes = sheet_additional_info;
                    }
                }
            }
            // ผลการพิจารณาอยู่ในกลุ่ม Cancelled
            // ไม่มีข้อมูลเพียงพอสำหรับการนำข้อมูล Cancel คำขอ	Error, ผิดพลาด ข้อมูลไม่เพียงพอสำหรับการนำข้อมูล Cancel คำขอ
            else if (cancelledResolution.includes(sheet_result)) {
                var resolution_code = get_active_result(sheet_result);
                if (resolution_code == "Error") {
                    // Close note ไม่มีใน template ข้อมูลที่จำเป็นจึงเหลือแค่ resolution code (ผลการพิจารณา)
                    add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด ข้อมูลไม่เพียงพอสำหรับการนำข้อมูล Cancel คำขอ");
                    error_log = error_log + 1;
                } else {
                    casetaskGr.u_resolution_code = resolution_code;
                    casetaskGr.close_notes = sheet_additional_info;
                    casetaskGr.state = 7; // Cancelled
                    casetaskGr.u_work_state = "ใบงานถูกยกเลิกโดยเจ้าหน้าที่";
                }
            } 
            else if(sheet_result == ""){
                add_log(source.sys_import_set, row_count, target_record, "Error", "โปรดระบุ ผลการพิจารณา");
                error_log = error_log + 1;
            }
            else {
                add_log(source.sys_import_set, row_count, target_record, "Error", "รูปแบบข้อมูล ผลการพิจารณา ไม่ถูกต้อง");
                error_log = error_log + 1;
            }

            // ภาระหนี้รวม (บาท)
            // ข้อมูลที่ใส่มาต้องเป็นตัวเงินเท่านั้น
            // if (check_money_valid(sheet_total_debt)) {
            //     casetaskGr.u_debt_burden = set_money_object(sheet_total_debt);
            // } else if (sheet_total_debt == "") {
            //     add_log(source.sys_import_set, row_count, target_record, "Error", "โปรดระบุ ภาระหนี้รวม (บาท)");
            //     error_log = error_log + 1;
            // } else {
            //     add_log(source.sys_import_set, row_count, target_record, "Error", "รูปแบบข้อมูล ภาระหนี้รวม (บาท) ไม่ถูกต้อง");
            //     error_log = error_log + 1;

            // }
            // เงินต้น (บาท)
            // เป็น Empty	Error, โปรดระบุ "เงินต้น (บาท)" 
            // มีค่าแต่ไม่ถูก format เช่น -1 หรือ Error Cast String to Decimal	Error, รูปแบบข้อมูล "เงินต้น (บาท)"  ไม่ถูกต้อง
            if (check_money_valid(sheet_principal)) {
                casetaskGr.u_principle = set_money_object(sheet_principal);
            } else if (sheet_principal == "") {
                add_log(source.sys_import_set, row_count, target_record, "Error", "โปรดระบุ เงินต้น (บาท)");
                error_log = error_log + 1;
            } else {
                add_log(source.sys_import_set, row_count, target_record, "Error", "รูปแบบข้อมูล เงินต้น (บาท) ไม่ถูกต้อง");
                error_log = error_log + 1;
            }

            // ภาระหนี้ที่ตกลงชำระ (บาท)
            // มีค่าแต่ไม่ถูก format เช่น -1 หรือ Error Cast String to Decimal	Error, รูปแบบข้อมูล ภาระหนี้ที่ตกลงชำระ ไม่ถูกต้อง
            // if (sheet_result == "ได้ข้อสรุปกับลูกค้า") {
            //     if (check_money_valid(sheet_debt_agreed)) {
            //         casetaskGr.u_debt_confirm = set_money_object(sheet_debt_agreed);
            //     } else if (sheet_debt_agreed == "") {
            //         //add code
            //         add_log(source.sys_import_set, row_count, target_record, "Error", "โปรดระบุ ภาระหนี้ที่ตกลงชำระ (บาท)");
            //         error_log = error_log + 1;
            //     } else {
            //         add_log(source.sys_import_set, row_count, target_record, "Error", "รูปแบบข้อมูล ภาระหนี้ที่ตกลงชำระ (บาท) ไม่ถูกต้อง");
            //         error_log = error_log + 1;

            //     }
            // }
            // จำนวนงวดที่ชำระ (เดือน)
            // ถ้าผลการพิจารณาเป็น ได้ข้อสรุปกับลูกค้า จำนวนงวดที่ต้องชำระ ห้ามเป็นค่าว่าง
            // ข้อมูลที่ใส่มาต้องเป็นตัวเลขจำนวนเต็มบวกเท่านั้น
            // if (sheet_result == "ได้ข้อสรุปกับลูกค้า") {
            //     if (sheet_no_payment == "") {
            //         add_log(source.sys_import_set, row_count, target_record, "Error", "โปรดระบุ จำนวนงวดที่ต้องชำระ (เดือน)");
            //         error_log = error_log + 1;
            //     } else if (validateNumber(sheet_no_payment)) {
            //         casetaskGr.u_installment = parseInt(sheet_no_payment);
            //     } else {
            //         add_log(source.sys_import_set, row_count, target_record, "Error", "รูปแบบ จำนวนงวดที่ชำระ (เดือน) ไม่ถูกต้อง");
            //         error_log = error_log + 1;
            //     }
            // }

            // ค่างวดต่อเดือน (บาท)
            // ถ้าผลการพิจารณาเป็น ได้ข้อสรุปกับลูกค้า ค่างวดต่อเดือน ห้ามเป็นค่าว่าง
            // ข้อมูลที่ใส่มาต้องเป็นตัวเงินเท่านั้น
            // if (sheet_result == "ได้ข้อสรุปกับลูกค้า") {
            //     if (check_money_valid(sheet_monthly)) {
            //         casetaskGr.u_amount_month = set_money_object(sheet_monthly);
            //     } else if (sheet_monthly == "") {
            //         add_log(source.sys_import_set, row_count, target_record, "Error", "โปรดระบุ ค่างวดต่อเดือน (บาท)");
            //         error_log = error_log + 1;
            //     } else {
            //         add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด ค่างวดต่อเดือน (บาท)");
            //         error_log = error_log + 1;
            //     }
            // }
            // รายงาน RDT
            // if (check_date_valid(sheet_rtd_report)) {
            //     casetaskGr.u_rdt = set_date_object(sheet_rtd_report);
            // } else if (sheet_rtd_report == "") {
            //     //add code
            // } else {
            //     add_log(source.sys_import_set, row_count, target_record, "Error", "รูปแบบข้อมูล วันที่ทำสัญญา  ไม่ถูกต้อง โปรดระบุใน format DD-MM-YYYY");
            //     error_log = error_log + 1;
            // }
            // วันที่ทำสัญญา
            // มีค่าแต่ไม่ถูก Error Cast String to Date	Error, รูปแบบข้อมูล "วันที่ทำสัญญา"  ไม่ถูกต้อง โปรดระบุใน format DD-MM-YYYY
            if (check_date_valid(sheet_contract_date)) {
                casetaskGr.u_drd = set_date_object(sheet_contract_date);
            } else if (sheet_contract_date == "") {
                //add code
            } else {
                add_log(source.sys_import_set, row_count, target_record, "Error", "รูปแบบข้อมูล วันที่ทำสัญญา  ไม่ถูกต้อง โปรดระบุใน format DD-MM-YYYY");
                error_log = error_log + 1;
            }

            // ข้อมูลที่ต้องการแจ้ง ธปท. เพิ่มเติม (ถ้ามี)
            // Length ยาวเกิน	Error, โปรดระบุ "ข้อมูลที่ต้องการแจ้ง ธปท. เพิ่มเติม (ถ้ามี)" ให้ขนาดไม่เกิน 4096
            if (sheet_additional_info.length > 4096) {
                add_log(source.sys_import_set, row_count, target_record, "Error", "ข้อมูลที่ต้องการแจ้ง ธปท. เพิ่มเติม (ถ้ามี) ยาวเกิน");
                error_log = error_log + 1;
            } else if (sheet_additional_info) {
                casetaskGr.close_notes = sheet_additional_info;
            }
            //วันที่รับคำขอ
            // ไม่เป็น Empty
            // มีค่าแต่ไม่ถูก Error Cast String to Date	
            // ถูกต้องตาม Format แต่วันที่เกินค่าวันปัจจุบัน	Error, ไม่สามารถกำหนดวันที่รับคำขอเป็นวันในอนาคตได้
            if (check_date_valid(sheet_take_request)) {
                //ตรวจสอบว่าวันที่รับคำขอเกินวันปัจจุบันหรือไม่
                if (checkDateMoreThanToday(set_date_object(sheet_take_request))) {
                    add_log(source.sys_import_set, row_count, target_record, "Error", "ไม่สามารถกำหนดวันที่รับคำขอเป็นวันในอนาคต");
                    error_log = error_log + 1;
                } else {
                    casetaskGr.opened_at = set_date_object(sheet_take_request);
                }
            } else if (sheet_take_request == "") {
                add_log(source.sys_import_set, row_count, target_record, "Error", "โปรดระบุ วันที่รับคำขอ");
                error_log = error_log + 1;
            } else {
                add_log(source.sys_import_set, row_count, target_record, "Error", "รูปแบบข้อมูล วันที่รับคำขอ ไม่ถูกต้อง โปรดระบุใน format DD-MM-YYYY");
                error_log = error_log + 1;
            }


            if (error_log == 0) {
                casetaskGr.insert();
                var case_number = casetaskGr.number;
                add_log(source.sys_import_set, row_count, case_number, "Insert", "สำเร็จ");
            } else {
                error_log_all = 1;
                add_log(source.sys_import_set, row_count, target_record, "Error", "Summary: ผิดพลาด Insert ไม่สำเร็จ มีข้อผิดพลาด " + String(error_log) + " ข้อ");
            }

        }
        // เรียกกการ log ว่าเป็น error หรือ complete หรือ complete with error ด้วยโค้ด
        var import_state = "";
        var log_total = 0;
        var log_error_tota = 0;
        var selectGrRecord = new GlideRecord("x_baot_debt_sett_0_import_excel_logs");
        selectGrRecord.addQuery("importsetref", source.sys_import_set);
        selectGrRecord.query();
        log_total = selectGrRecord.getRowCount();

        var erorrGrRecord = new GlideRecord("x_baot_debt_sett_0_import_excel_logs");
        erorrGrRecord.addQuery("importsetref", source.sys_import_set);
        erorrGrRecord.addQuery("state", 'ERROR');
        erorrGrRecord.query();
        log_error_total = erorrGrRecord.getRowCount();

        var row_table_importattachment = new GlideRecord('x_baot_debt_sett_0_import_excel_attachments');
        row_table_importattachment.addQuery("sys_id", importExcelAttachmentSysId);
        row_table_importattachment.query();
        if (row_table_importattachment.next()) {

            row_table_importattachment.setValue("importset_link", source.sys_import_set);
            gs.info("importset_link" + source.sys_import_set);
            gs.info("[Excel Bulk Upload]: log_total" + log_total + " log_error_total:" + log_error_total);
            if (log_total == log_error_total) {
                // row_table_importattachment.setValue("state", "5"); // Error == 5
                gs.info("[Excel Bulk Upload state]: Error");
            } else if (log_error_total == 0) {
                // row_table_importattachment.setValue("state", "3"); // Completed == 3
                gs.info("[Excel Bulk Upload state]: Completed");
            } else {
                // row_table_importattachment.setValue("state", "4"); // Completed with error == 4
                gs.info("[Excel Bulk Upload state]: Completed with error");
            }
            var now_time = new GlideDateTime();
            var last_time = new GlideDateTime(row_table_importattachment.sys_created_on);
            var durationsmilli = now_time.getNumericValue() - last_time.getNumericValue();

            row_table_importattachment.loaded_time = now_time;
            row_table_importattachment.load_runtime = new GlideDuration(durationsmilli);
            row_table_importattachment.update();

        } else {
            gs.info("[Excel Bulk Upload] Record row_table_importattachment not found!");
        }

        if (error_log_all != 0) {
            error = 1;
        }

    } catch (error) {
        gs.info("[Excel Bulk Upload] error: " + error);


        var sysid_table_importattachment = importExcelAttachmentSysId;
        var row_table_importattachment = new GlideRecord('x_baot_debt_sett_0_import_excel_attachments');
        row_table_importattachment.addQuery("sys_id", sysid_table_importattachment);
        row_table_importattachment.query();
        if (row_table_importattachment.next()) {
            row_table_importattachment.setValue("importset_link", source.sys_import_set);
            row_table_importattachment.setValue("state", "5");
            gs.info("row_table_importattachment.setValue(5)");

            var now_time = new GlideDateTime();
            var last_time = new GlideDateTime(row_table_importattachment.sys_created_on);
            var durationsmilli = now_time.getNumericValue() - last_time.getNumericValue();

            row_table_importattachment.loaded_time = now_time;
            row_table_importattachment.load_runtime = new GlideDuration(durationsmilli);
            row_table_importattachment.update();
        } else {
            gs.info("[Excel Bulk Upload Catch Error] Record row_table_importattachment not found!");
        }
    }


})(source, map, log, target);


function formatDate(inputDate) {
    // Convert date string from dd-mm-yyyy to yyyy-mm-dd
    // return formatted date string if valid, empty string if invalid
    var yearFirstPattern = /^(\d{4})[-\/](\d{2})[-\/](\d{2})$/;
    var dayFirstPattern = /^(\d{2})[-\/](\d{2})[-\/](\d{4})$/;
    var formattedDate;
    if (yearFirstPattern.test(inputDate)) {
        var match = inputDate.match(yearFirstPattern);
        formattedDate = match[3] + '-' + match[2] + '-' + match[1];
    } else if (dayFirstPattern.test(inputDate)) {
        var match1 = inputDate.match(dayFirstPattern);
        formattedDate = match1[1] + '-' + match1[2] + '-' + match1[3];
    } else {
        formattedDate = "";
    }
    return formattedDate;
}

function add_log(importno, row, target_record, state, message) {
    gs.info("[Excel Bulk Upload]" + " " + "[Row " + row + " ]" + "[State " + state + " ]" + " " + message + "[importno " + importno + " ]");

    var logrow = new GlideRecord("x_baot_debt_sett_0_import_excel_logs");
    var userSysId = gs.getUserID();
    logrow.initialize();
    logrow.importsetref = importno;
    logrow.row = parseInt(row) + 1;
    logrow.state = state;
    logrow.target_record = target_record;
    logrow.message = message;
    logrow.insert();
}

function get_select_valid_active(table_select, choice_col, choice_select) {
    var selectGr = new GlideRecord(table_select);
    selectGr.addQuery(choice_col, choice_select);
    selectGr.setLimit(1);
    selectGr.query();
    if (selectGr.next()) {
        if (selectGr.u_active == true) {
            return selectGr.getUniqueValue();
        }
    } else {
        return "Error";
    }
}

function check_datetime_valid(date_input) {
    // Check if date is a valid date format (dd-mm-yyyy hh:mm)
    // return true if valid, false if invalid
    var regex = /^(\d{2})-(\d{2})-(\d{4}) (\d{2}):(\d{2})$/;
    var match2 = false;
    try {
        match2 = date_input.match(regex);
    } catch (err) {
        match2 = false;
    }

    if (!match2) {
        return false;
    }

    var day = parseInt(match2[1], 10);
    var month = parseInt(match2[2], 10);
    var year = parseInt(match2[3], 10);
    var hour = parseInt(match2[4], 10);
    var minute = parseInt(match2[5], 10);

    if (day < 1 || day > 31 || month < 1 || month > 12 || hour < 0 || hour > 23 || minute < 0 || minute > 59) {
        return false;
    }

    var daysInMonth = [31, (year % 4 == 0 && year % 100 !== 0) || (year % 400 == 0) ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    if (day > daysInMonth[month - 1]) {
        return false;
    }

    return true;
}

function set_timezone_object(datetime_string_input) {
    // Convert date string from dd-mm-yyyy hh:mm to yyyy-mm-dd hh:mm:ss
    var datePart = datetime_string_input.split(" ")[0];
    var timePart = datetime_string_input.split(" ")[1];
    var dateArray = datePart.split("-");

    var formattedDate = dateArray[2] + "-" + dateArray[1] + "-" + dateArray[0] + " " + timePart + ":00";

    var gdt1 = new GlideDateTime(formattedDate);
    gdt1.addSeconds(-25200);


    return gdt1;
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

function check_date_valid(date_input) {
    // Check if date is a valid date format (dd-mm-yyyy)
    // return true if valid, false if invalid
    var regex = /^(\d{2})-(\d{2})-(\d{4})$/;
    var match2 = false;
    try {
        match2 = date_input.match(regex);
    } catch (err) {
        match2 = false;
    }
    if (!match2) {
        return false;
    }

    const day = parseInt(match2[1], 10);
    const month = parseInt(match2[2], 10);
    const year = parseInt(match2[3], 10);

    if (day < 1 || day > 31 || month < 1 || month > 12) {
        return false;
    }

    const daysInMonth = [31, (year % 4 === 0 && year % 100 !== 0) || (year % 400 === 0) ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    if (day > daysInMonth[month - 1]) {
        return false;
    }

    return true;
}

function get_select_valid_active_check_fair(table_select, choice_col, choice_select, debtFair) {
    var grFair = new GlideRecord("u_project_debt");
    grFair.addEncodedQuery("u_name=" + debtFair);
    grFair.query();
    if (grFair.next()) {
        var selectGr = new GlideRecord(table_select);
        selectGr.addQuery(choice_col, choice_select);
        selectGr.addQuery("u_debt_fair", grFair.sys_id);
        selectGr.setLimit(1);
        selectGr.query();
        if (selectGr.next()) {
            if (selectGr.u_active == true) {
                return selectGr.getUniqueValue();
            }
        } else {
            return "Error";
        }
    }
}

function check_money_valid(amount) {
    // Check if amount is a valid money format (THB) with comma and 2 decimal places
    var regex = /^(฿)?(\d{1,3}(,\d{3})*|\d+)(\.\d{1,2})?$/;
    return regex.test(amount);
}

function set_money_object(money_input) {
    var x = money_input.replace("฿", "");
    var y = x.replace(",", "");
    return "THB;" + y;
}

function getRef(table_name, field_name, value) {
    var gr = new GlideRecord(table_name);
    gr.addQuery(field_name, value);
    gr.query();
    if (gr.next()) {
        return gr.sys_id;
    } else {
        return "Error";
    }
}

function get_user_valid_active(table_select, choice_col, choice_select) {
    var selectGr = new GlideRecord(table_select);
    selectGr.addQuery(choice_col, choice_select);
    selectGr.setLimit(1);
    selectGr.query();
    if (selectGr.next()) {
        return selectGr.getUniqueValue();
    } else {
        return "Error";
    }
}

function set_date_object(date_string_input) {
    // Convert date string from dd-mm-yyyy to yyyy-mm-dd
    const regex = /^(\d{2})-(\d{2})-(\d{4})$/;
    const match = date_string_input.match(regex);
    const day = parseInt(match[1], 10);
    const month = parseInt(match[2], 10);
    const year = parseInt(match[3], 10);

    return year + "-" + month + "-" + day;
}

function validateName(name) {
    //english, thai, space, hyphen only, not empty
    var NameRegex = /^[A-Za-zก-๙'-]+(?: [A-Za-zก-๙'-]+)*$/;
    return NameRegex.test(name) && name.length >= 2 && name.length <= 100;
}

function validateString(str) {
    //english, thai, space only, not empty
    var Regex = /^[A-Za-zก-๙\s0-9]+$/;
    return Regex.test(str) && str.length >= 2 && str.length <= 100;
}

function validateNumber(str) {
    //number (String) only, not empty
    if (typeof str !== "string") return false;
    var Regex = /^\d+$/;
    return Regex.test(str) && str.length >= 1 && str.length <= 100;
}

function validatePhone(str) {
    //numberE164 only eg. +66812345678 , not empty
    //+66812345678 true
    //0812345678 false
    var Regex = /^\+66(2\d{7}|[689]\d{8})$/;
    return Regex.test(str);
}

function validateEmail(email) {
    //email only, not empty
    var EmailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
    return EmailRegex.test(email) && email.length >= 2 && email.length <= 100;
}

function checkDateMoreThanToday(date_input) {
    //campare data YYYY-MM-DD input now date and date return true if date input <= now date
    var now = new GlideDateTime();
    var date = new GlideDateTime(date_input);
    if (date.compareTo(now) <= 0) {
        return false;
    } else {
        return true;
    }
}

function isNotMember(userSysId, groupSysId) {
    var memberOfGroup = true;
    var gr = new GlideRecord('sys_user_grmember');
    gr.addQuery('user', userSysId);
    gr.addQuery('group', groupSysId);
    gr.query();
    if (gr.next()) {
        memberOfGroup = false;
    }

    return memberOfGroup;
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

function check_select_valid_activePipe(table_select, enCodeQuery) {
    //gs.info('Select > ' + table_select + ' ' + choice_col + ' ' + choice_select);
    var selectGr = new GlideRecord(table_select);
    selectGr.addEncodedQuery(enCodeQuery);
    gs.info('Select > ' + table_select + ' ' + enCodeQuery);
    selectGr.setLimit(1);
    selectGr.query();
    if (selectGr.next()) {
        return true;
    } else {
        return false;
    }
}