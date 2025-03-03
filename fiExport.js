
function walkinFIbulkUpload(queryTable, query, datee, fairquery, external) {
	var gr = new GlideRecord(queryTable);
	gr.addEncodedQuery(query + datee + fairquery + external);
	gr.query();

	var records = [];
	while (gr.next()) {
		var recordObj = {};
        // หมายเลข Case Task ok
		recordObj.number = gr.getValue('number');
        // เรื่อง ok
		recordObj.short_description = gr.getValue('short_description');
        // ผู้ขอรับบริการ u_walk_service_requester
		recordObj.u_consumer = gr.getDisplayValue('u_name_consumer');
        // Due date ok
		recordObj.due_date = gr.getDisplayValue('due_date');
        // SLA Due date ok
		recordObj.u_sla_due_date = gr.getDisplayValue('u_sla_due_date');
        // กลุ่มที่รับผิดชอบ ok
		recordObj.assignment_group = gr.getDisplayValue('assignment_group');
        // ผู้รับมอบหมาย ok
		recordObj.assigned_to = gr.getDisplayValue('assigned_to.email');
        // ขอรับบริการในนาม u_walkin_receive_service
		recordObj.u_receive_service = gr.u_receive_service.getDisplayValue();
        // ผู้ติดต่อ pending
		recordObj.u_exact_consumer = gr.getDisplayValue('u_name_exact_consumer');
        // หมายเลข Case ok
		recordObj.u_display_number = gr.getDisplayValue('u_display_number');
        // ประเภทเลขที่ยืนยันตัวตน (นิติบุคคล) ok
		recordObj.u_type_juristic = gr.u_type_juristic.getDisplayValue();
        // ประเภทเลขที่ยืนยันตัวตน ok
		recordObj.u_type_iden = gr.getDisplayValue('u_type_iden');
        // เลขที่ยืนยันตัวตน u_walkin_identification_number
		recordObj.u_iden_number = gr.getDisplayValue('u_iden_number');
        // เลขที่ยืนยันตัวตน (นิติบุคคล) u_identification_number_juristic
        recordObj.u_number_juristic = gr.u_number_juristic.getDisplayValue();
        // ชื่อนิติบุคคล u_walkin_juristic_name
		recordObj.u_juristic_name = gr.u_juristic_name.getDisplayValue();
        // หมายเลขโทรศัพท์ u_walkin_phone
		recordObj.u_mobile = gr.getDisplayValue('u_mobile');
        // หมายเลขโทรศัพท์ (สำรอง) ok
		recordObj.u_secondary_phone = gr.u_secondary_phone.getDisplayValue();
        // อีเมล u_walkin_email
		recordObj.u_mail = gr.getDisplayValue('u_mail');
        // อีเมล (สำรอง) ok
		recordObj.u_secondary_email = gr.getDisplayValue('u_secondary_email');
        // จังหวัด (ที่อยู่ลูกหนี้) ok
		recordObj.u_province = gr.getDisplayValue('u_province');
        // ผู้ให้บริการ ok
		recordObj.u_provider = gr.getDisplayValue('u_provider');
        // โครงการแก้หนี้ ok
		recordObj.u_debt_project = gr.getDisplayValue('u_debt_project');
        // แนวทางที่ต้องการเสนอให้ผู้บริการทางการเงินพิจารณา ok
		recordObj.u_offer_provider = gr.getDisplayValue('u_offer_provider');
        // เหตุผลหรือคำอธิบายประกอบคำขอ ok
		recordObj.u_reason = gr.getDisplayValue('u_reason');
        // Case ที่เกี่ยวข้อง ok
		recordObj.u_major_case = gr.getDisplayValue('u_major_case');
        // รายละเอียดหนี้ ok
		recordObj.u_debt_detail = gr.getDisplayValue('u_debt_detail');
        // เหตุผลที่ขอดำเนินเรื่องต่อ ok
		recordObj.u_reason_reopen_choice = gr.getDisplayValue('u_reason_reopen_choice');
        // รายละเอียดที่ขอดำเนินเรื่องต่อ ok
		recordObj.u_reason_reopen = gr.getDisplayValue('u_reason_reopen');
        // สถานะการดำเนินงาน ok
		recordObj.u_work_state = gr.getValue('u_work_state');
        // ประเภทเลขที่ยืนยันตัวตน ok
		recordObj.u_identidifier_type = gr.u_iden_number.getDisplayValue();
        // ไม่ถูกใช้ในการ map
		recordObj.u_phone_numer = gr.u_mobile.getDisplayValue();
        // สถานะใบงาน ok
		recordObj.state = gr.getDisplayValue('state');
        // เลขที่ทะเบียนรถ ok
		recordObj.u_number_car = gr.getValue('u_number_car');
        // จังหวัดที่จดทะเบียนรถ ok
		recordObj.u_province_car = gr.getDisplayValue('u_province_car');
        // เลขที่บัตร/ เลขที่สัญญา ok
		recordObj.u_number_contact = gr.getValue('u_number_contact');
        // วันที่เริ่มติดต่อลูกหนี้ ok
		recordObj.u_contact_debt = gr.getDisplayValue('u_contact_debt');
        // ผลการพิจารณา ok
		recordObj.u_resolution_code = gr.getDisplayValue('u_resolution_code');

		var grXBDS_0PTF = new GlideRecord('x_baot_debt_sett_0_product_to_fair');
		grXBDS_0PTF.addEncodedQuery("x_baot_debt_sett_0_debt_fair!=bd12cc231b5dbd1080bb844ee54bcb5c^ORx_baot_debt_sett_0_debt_fair=NULL^u_m2m_product_to_application.sys_id=" + gr.getValue('u_product'));
		grXBDS_0PTF.query();

		if (grXBDS_0PTF.next()) {
            // ผลิตภัณฑ์ ok
			recordObj.u_product = grXBDS_0PTF.getDisplayValue("u_m2m_product_to_application") + " | " + gr.getDisplayValue('u_debt_project');
		}

		var grXBDS_0DTFStateDept = new GlideRecord('x_baot_debt_sett_0_debt_to_fair');
		grXBDS_0DTFStateDept.addEncodedQuery("x_baot_debt_sett_0_debt_fair!=a86920c51b2cb1106c24dcace54bcb8a^ORx_baot_debt_sett_0_debt_fair=NULL^x_baot_debt_sett_0_debt_fair!=6904000d1bacf11080bb844ee54bcb95^ORx_baot_debt_sett_0_debt_fair=NULL^x_baot_debt_sett_0_debt_fair!=bd12cc231b5dbd1080bb844ee54bcb5c^ORx_baot_debt_sett_0_debt_fair=NULL^x_baot_debt_sett_0_debt_fair!=cd924c631b5dbd1080bb844ee54bcbbc^ORx_baot_debt_sett_0_debt_fair=NULL^x_baot_debt_sett_0_debt_fair!=dd944f0c1b34b91080bb844ee54bcbf9^ORx_baot_debt_sett_0_debt_fair=NULL^u_m2m_debt_status_to_application.sys_id=" + gr.getValue("u_state_debt"));
		grXBDS_0DTFStateDept.query();

		if (grXBDS_0DTFStateDept.next()) {
            // สถานะบัญชี ok
			recordObj.u_state_debt = grXBDS_0DTFStateDept.getDisplayValue("u_m2m_debt_status_to_application") + " | " + gr.getDisplayValue('u_debt_project');
		}
        // สถานะใบงาน ok
		recordObj.state = gr.getDisplayValue('state');
        // เลขที่ที่ทะเบียนรถ ok
		recordObj.u_number_car = gr.getValue('u_number_car');
        // จังหวัดที่จดทะเบียนรถ ok
		recordObj.u_province_car = gr.getDisplayValue('u_province_car');
        // เลขที่บัตร/ เลขที่สัญญา ok
		recordObj.u_number_contact = gr.getValue('u_number_contact');
        // วันที่เริ่มติดต่อลูกหนี้ ok
		recordObj.u_contact_debt = gr.getDisplayValue('u_contact_debt');
        // ผลการพิจารณา ok
		recordObj.u_resolution_code = gr.getDisplayValue('u_resolution_code');

		var grUGD = new GlideRecord('u_glideline_debtor');
		grUGD.addEncodedQuery("sys_id=" + gr.getValue("u_glideline_debt"));
		grUGD.query();
		if (grUGD.next()) {
            // แนวทางการช่วยเหลือลูกหนี้ ok
			recordObj.u_glideline_debt = grUGD.getDisplayValue('u_name') + " | " + grUGD.getDisplayValue('u_debt_fair');
		}

		var grUUH = new GlideRecord('u_unable_help');
		grUUH.addEncodedQuery("sys_id=" + gr.getValue("u_reason_unable_help"));
		grUUH.query();
		if (grUUH.next()) {
            // เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้ ok
			recordObj.u_reason_unable_help = grUUH.getDisplayValue('u_name') + " | " + grUUH.getDisplayValue('u_debt_fair');
		}
        // ภาระหนี้รวม (บาท) ok
		recordObj.u_debt_burden = gr.getDisplayValue('u_debt_burden');
        // เงินต้น (บาท) ok
		recordObj.u_principle = gr.getDisplayValue('u_principle');
        // ภาระหนี้ที่ตกลงชำระ (บาท) ok
		recordObj.u_debt_confirm = gr.getDisplayValue('u_debt_confirm');
        // จำนวนงวดที่ชำระ (เดือน) ok
		recordObj.u_installment = gr.getDisplayValue('u_installment');
        // ค่างวดต่อเดือน (บาท) ok
		recordObj.u_amount_month = gr.getDisplayValue('u_amount_month');
        // วันที่ทำสัญญา ok
		recordObj.u_drd = gr.getDisplayValue('u_drd');
        // รายงาย RDT ok 
		recordObj.u_rdt = gr.getDisplayValue('u_rdt');
        // ข้อมูลที่ต้องการแจ้ง ธปท. เพิ่มเติม (ถ้ามี) close_notes
		recordObj.u_notify_debtor = gr.getDisplayValue('u_notify_debtor');
        // ไม่ถูกใช้ในการ map
		recordObj.closed_at = gr.getDisplayValue('closed_at');
        // จำนวนการ Re-open ok
		recordObj.u_re_open_count = gr.getDisplayValue('u_re_open_count');
        // วันที่รับคำขอ ok
		recordObj.opened_at = gr.getDisplayValue('opened_at');

		records.push(recordObj);
	}
	return JSON.stringify(records);
}




// ชื่อฟิลจริงที่ใช้ในการเอาเช้า record
var nameField = ["number","short_description","u_consumer","due_date","u_sla_due_date","assignment_group","assigned_to","u_receive_service","u_exact_consumer","u_display_number","u_type_juristic","u_type_iden","u_iden_number","u_number_juristic","u_juristic_name","u_mobile","u_secondary_phone","u_mail","u_secondary_email","u_province","u_provider","u_debt_project","u_offer_provider","u_reason","u_major_case","u_debt_detail","u_reason_reopen_choice","u_reason_reopen","u_work_state","u_identidifier_type","u_phone_numer","state","u_number_car","u_province_car","u_number_contact","u_contact_debt","u_resolution_code","u_product","u_state_debt","u_glideline_debt","u_reason_unable_help","u_debt_burden","u_principle","u_debt_confirm","u_installment","u_amount_month","u_drd","u_rdt","u_notify_debtor","closed_at","u_re_open_count","opened_at"];

var mapWord = [
    {
      "หมายเลข Case Task": "number"
    },
    {
      "เรื่อง": "short_description"
    },
    {
      "สถานะการดำเนินงาน": "u_work_state"
    },
    {
      "SLA Due date": "u_sla_due_date"
    },
    {
      "กลุ่มที่รับผิดชอบ": "assignment_group"
    },
    {
      "Due date": "due_date"
    },
    {
      "ผู้รับมอบหมาย": "assigned_to"
    },
    {
      "หมายเลข Case": "u_display_number"
    },
    {
      "ขอรับบริการในนาม": "u_receive_service"
    },
    {
      "ผู้ติดต่อ": "u_exact_consumer"
    },
    {
      "ผู้ขอรับบริการ": "u_consumer"
    },
    {
      "ประเภทเลขที่ยืนยันตัวตน": "u_type_iden"
    },
    {
      "ประเภทเลขที่ยืนยันตัวตน (นิติบุคคล)": "u_type_juristic"
    },
    {
      "เลขที่ยืนยันตัวตน": "u_iden_number"
    },
    {
      "เลขที่ยืนยันตัวตน (นิติบุคคล)": "u_number_juristic"
    },
    {
      "หมายเลขโทรศัพท์": "u_mobile"
    },
    {
      "อีเมล": "u_mail"
    },
    {
      "หมายเลขโทรศัพท์ (สำรอง)": "u_secondary_phone"
    },
    {
      "อีเมล (สำรอง)": "u_secondary_email"
    },
    {
      "จังหวัด (ที่อยู่ลูกหนี้)": "u_province"
    },
    {
      "ชื่อนิติบุคคล": "u_juristic_name"
    },
    {
      "ผู้ให้บริการ": "u_provider"
    },
    {
      "ผลิตภัณฑ์": "u_product"
    },
    {
      "สถานะบัญชี": "u_state_debt"
    },
    {
      "โครงการแก้หนี้": "u_debt_project"
    },
    {
      "เลขที่ทะเบียนรถ": "u_number_car"
    },
    {
      "จังหวัดที่จดทะเบียนรถ": "u_province_car"
    },
    {
      "แนวทางที่ต้องการเสนอให้ผู้บริการทางการเงินพิจารณา": "u_offer_provider"
    },
    {
      "เหตุผลหรือคำอธิบายประกอบคำขอ": "u_reason"
    },
    {
      "Case ที่เกี่ยวข้อง": "u_major_case"
    },
    {
      "Ref Case Task": ""
    },
    {
      "เลขที่บัตร/ เลขที่สัญญา": "u_number_contact"
    },
    {
      "วันที่เริ่มติดต่อลูกหนี้": "u_contact_debt"
    },
    {
      "ผลการพิจารณา": "u_resolution_code"
    },
    {
      "แนวทางการช่วยเหลือลูกหนี้": "u_glideline_debt"
    },
    {
      "เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้": "u_reason_unable_help"
    },
    {
      "ภาระหนี้รวม (บาท)": "u_debt_burden"
    },
    {
      "เงินต้น (บาท)": "u_principle"
    },
    {
      "ภาระหนี้ที่ตกลงชำระ (บาท)": "u_debt_confirm"
    },
    {
      "จำนวนงวดที่ชำระ (เดือน)": "u_installment"
    },
    {
      "ค่างวดต่อเดือน (บาท)": "u_amount_month"
    },
    {
      "รายงาน RDT": "u_rdt"
    },
    {
      "วันที่ทำสัญญา": "u_drd"
    },
    {
      "ข้อมูลที่ต้องการแจ้ง ธปท. เพิ่มเติม (ถ้ามี)": "u_notify_debtor"
    },
    {
      "สถานะใบงาน": "state"
    },
    {
      "Close note": ""
    },
    {
      "จำนวนการ Re-open": "u_re_open_count"
    },
    {
      "เหตุผลที่ขอดำเนินเรื่องต่อ": "u_reason_reopen_choice"
    },
    {
      "รายละเอียดที่ขอดำเนินเรื่องต่อ": "u_reason_reopen"
    },
    {
      "วันที่รับคำขอ": "opened_at"
    }
  ];