(function runTransformScript(source, map, log, target) {


	var importExcelAttachmentSysId = "";
	try {
		gs.info("Start Transform Script : onCompletee");
		var targetTable = source.getTableName();
		var targetGR = new GlideRecord(targetTable);
		targetGR.addQuery('sys_import_set', source.sys_import_set);
		targetGR.query();
		var count = targetGR.getRowCount();
		var row_count = 0;
		var error_log = 0;
		var error_log_all = 0;
		var target_record = "-";
		var userSysId_uploadby = "";
		var debt_project;
		var countNumber = 1;

		// update ข้อมูลด้วยตนเอง
		var carProduct = ["สินเชื่อเช่าซื้อรถยนต์ / Top-up / รถแลกเงินแบบโอนเล่ม", "สินเชื่อเช่าซื้อรถจักรยานยนต์ / Top-up / รถแลกเงินโอนเล่ม", "จำนำทะเบียนรถยนต์ (ไม่โอนเล่ม)", "จำนำทะเบียนรถจักรยานยนต์ (ไม่โอนเล่ม)", "สินเชื่อเช่าซื้อรถยนต์", "สินเชื่อเช่าซื้อรถจักรยานยนต์", "สินเชื่อจำนำทะเบียนรถยนต์", "สินเชื่อจำนำทะเบียนรถจักรยานยนต์"];


		var grSA = new GlideRecord('sys_attachment');
		grSA.addEncodedQuery("table_sys_idSTARTSWITHd4093f2747d3c250ab9eb5c8736d43da");
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

		gs.info("[Excel Bulk Upload]get user: " + userSysId_uploadby);

		var recordsLimit = gs.getProperty('x_baot_debt_sett_0.excelRecordsLimit');
		var countCondition = false;
		if (recordsLimit < 1) countCondition = count < 1;
		else countCondition = count < 1 || count > recordsLimit;

		if (countCondition) {
			var gr = new GlideRecord('sys_attachment');
			// gr.addQuery('table_name', 'x_baot_debt_sett_0_import_excel_attachments');
			grSA.addEncodedQuery("table_sys_idSTARTSWITHd4093f2747d3c250ab9eb5c8736d43da");
			gr.orderByDesc('sys_created_on');
			gr.query();
			if (gr.next()) {
				var sysid_table_importattachment = gr.getValue("file_name").split("-")[0];
				var row_table_importattachment = new GlideRecord('x_baot_debt_sett_0_import_excel_attachments');
				row_table_importattachment.addQuery('sys_id', sysid_table_importattachment);
				row_table_importattachment.query();
				if (row_table_importattachment.next()) {
					row_table_importattachment.state = 5;
					if (count < 1) {
						row_table_importattachment.u_file_error = 'ไฟล์ที่แนบไม่มีข้อมูล';
					} else if (recordsLimit > 0) {
						row_table_importattachment.u_file_error = 'ไฟล์ที่แนบต้องมีข้อมูลไม่เกิน ' + recordsLimit + ' records';
					} else {
						row_table_importattachment.u_file_error = 'ไฟล์ที่แนบมีข้อมูลไม่ถูกต้อง';
					}
					row_table_importattachment.update();

				}
			}


		} else {
			while (targetGR.next()) {
				row_count += 1;
				error_log = 0;
				target_record = "-";
				var sheet_case_id = targetGR.u_หมายเลข_case;

				// หมายเลข Case Task
				var sheet_case_task_id = targetGR.u_หมายเลข_case_task;
				// 
				var sheet_case_relate = targetGR.u_case_ท__เก__ยวข_อง;
				//case detail
				var sheet_case_detail = targetGR.u_หมายเลข_detail;
				// เรื่อง
				var sheet_title = targetGR.u_เร__อง;
				// สถานะการดำเนินงาน
				var sheet_status = targetGR.u_สถานะการดำเน_นงาน;
				// SLA Due date
				var sheet_sla_duedate = targetGR.u_sla_due_date;
				// กลุ่มที่รับผิดชอบ
				var sheet_responsible_group = targetGR.u_กล__มท__ร_บผ_ดชอบ;
				// Due date
				var sheet_duedate = targetGR.u_due_date;
				// ผู้รับมอบหมาย
				var sheet_assignee = targetGR.u_ผ__ร_บมอบหมาย;
				// ขอรับบริการในนาม
				var sheet_request_service_on_behalf_of = targetGR.u_ขอร_บบร_การในนาม;
				// ผู้ติดต่อ
				var sheet_contact = targetGR.u_ผ__ต_ดต_อ;
				// ผู้ขอรับบริการ
				var sheet_service_requester = targetGR.u_ผ__ขอร_บบร_การ;
				// ประเภทเลขที่ยืนยันตัวตน
				var sheet_type_identification_number = targetGR.u_ประเภทเลขท__ย_นย_นต_วตน;
				// ประเภทเลขที่ยืนยันตัวตน (นิติบุคคล)
				var sheet_type_identification_number_juristic = targetGR.u_ประเภทเลขท___ตน__น_ต_บ_คคล_;
				// เลขที่ยืนยันตัวตน
				var sheet_identification_number = targetGR.u_เลขท__ย_นย_นต_วตน;
				// เลขที่ยืนยันตัวตน (นิติบุคคล)
				var sheet_identification_number_juristic = targetGR.u_เลขท__ย_นย_น_ตน__น_ต_บ_คคล_;
				// หมายเลขโทรศัพท์
				var sheet_phone = targetGR.u_หมายเลขโทรศ_พท_;
				// อีเมล
				var sheet_email = targetGR.u_อ_เมล;
				// หมายเลขโทรศัพท์ (สำรอง)
				var sheet_phone_backup = targetGR.u_หมายเลขโทรศ_พท___สำรอง_;
				// อีเมล (สำรอง)
				var sheet_email_backup = targetGR.u_อ_เมล__สำรอง_;
				// จังหวัด (ที่อยู่ลูกหนี้)
				var sheet_province = targetGR.u_จ_งหว_ด__ท__อย__ล_กหน___;
				// ผู้ให้บริการ
				var sheet_provider = targetGR.u_ผ__ให_บร_การ;

				var sheet_product = targetGR.u_ผล_ตภ_ณฑ_;
				if (sheet_product == '' || sheet_product == null) {
					//    gs.info('sheet_guidelines_debtorssssssssssssllllllllllds');

				} else {
					sheet_product = targetGR.u_ผล_ตภ_ณฑ_.split("|")[0].trim();

				}

				var sheet_account_status = targetGR.u_สถานะบ_ญช_;
				if (sheet_account_status == '' || sheet_account_status == null) {
					//    gs.info('sheet_guidelines_debtorssssssssssssllllllllllds');

				} else {
					sheet_account_status = targetGR.u_สถานะบ_ญช_.split("|")[0].trim();

				}

				//โครงการแก้หนี้
				var sheet_project_debt = targetGR.u_โครงการแก_หน__;

				// เลขที่ทะเบียนรถ
				var sheet_vehicle_number = targetGR.u_เลขท__ทะเบ_ยนรถ;
				// จังหวัดที่จดทะเบียนรถ
				var sheet_vehicle_number_province = targetGR.u_จ_งหว_ดท__จดทะเบ_ยนรถ;
				// แนวทางที่ต้องการเสนอให้ผู้บริการทางการเงินพิจารณา
				var sheet_guidelines_provider = targetGR.u_แนวทางท__ต_อ_การเง_นพ_จารณา;
				// เหตุผลหรือคำอธิบายประกอบคำขอ
				var sheet_reason = targetGR.u_เหต_ผลหร_อคำอธ_บายประกอบคำขอ;
				// // Case ที่เกี่ยวข้อง
				var sheet_case_relate_2 = targetGR.u_case_ท__เก__ยวข_อง;
				// Ref Case Task
				var sheet_case_ref = targetGR.u_ref_case_task;
				// เลขที่บัตร/ เลขที่สัญญา
				var sheet_contract_number = targetGR.u_เลขท__บ_ตร__เลขท__ส_ญญา;
				// วันที่เริ่มติดต่อลูกหนี้
				var sheet_date_of_start = targetGR.u_ว_นท__เร__มต_ดต_อล_กหน__;
				// ผลการพิจารณา
				var sheet_result = targetGR.u_ผลการพ_จารณา;
				// แนวทางการช่วยเหลือลูกหนี้
				var sheet_guidelines_debtors = targetGR.u_แนวทางการช_วยเหล_อล_กหน__;
				if (sheet_guidelines_debtors == '' || sheet_guidelines_debtors == null) {
					//gs.info('sheet_guidelines_debtorssssssssssssllllllllllds');
				} else {
					//  gs.info('sheet_guidelines_debtorssssssssseeeeeeeeeeeeeeeedlds');
					sheet_guidelines_debtors = targetGR.u_แนวทางการช_วยเหล_อล_กหน__.split("|")[0].trim();
				}

				var sheet_reason_not_help = targetGR.u_เหต_ผลท__ไม__หล_อล_กหน__ได_;
				if (sheet_reason_not_help == '' || sheet_reason_not_help == null) {
					//    gs.info('sheet_guidelines_debtorssssssssssssllllllllllds');

				} else {
					//  gs.info('sheet_guidelines_debtorssssssssseeeeeeeeeeeeeeeedlds');
					sheet_reason_not_help = targetGR.u_เหต_ผลท__ไม__หล_อล_กหน__ได_.split("|")[0].trim();
				}

				var sheet_total_debt = targetGR.u_ภาระหน__รวม__บาท_;
				// เงินต้น (บาท)
				var sheet_principal = targetGR.u_เง_นต_น__บาท_;
				// ภาระหนี้ที่ตกลงชำระ (บาท)
				var sheet_debt_agreed = targetGR.u_ภาระหน__ท__ตกลงชำระ__บาท_;
				// จำนวนงวดที่ชำระ (เดือน)
				var sheet_no_payment = targetGR.u_จำนวนงวดท__ชำระ__เด_อน_;
				// ค่างวดต่อเดือน (บาท)
				var sheet_monthly = targetGR.u_ค_างวดต_อเด_อน__บาท_;
				// รายงาน RDT
				var sheet_rdt_report = targetGR.u_รายงาน_rdt;
				// รายงาน DRD
				var sheet_drd_report = targetGR.u_รายงาน_drd ? targetGR.u_รายงาน_drd : targetGR.u_ว_นท__ทำส_ญญา;
				// ข้อมูลที่ต้องการแจ้ง ธปท. เพิ่มเติม (ถ้ามี)
				var sheet_additional_info = targetGR.u_ข_อม_ลท__ต_อ__มเต_ม__ถ_าม__;

				var sheet_state = targetGR.u_สถานะใบงาน;
				var sheet_cancel_state = targetGR.u_ประเภทการแก_ไข;
				var sheet_close = targetGR.u_close_note;

				//เพิ่มการ update ของตัว FI ว่าเข้ามาเป็นอันนี้แล้วต้องเรียก table ของ FI
				if (sheet_case_task_id && sheet_case_task_id.substring(0, 3) == "DET") { // If excel uploaded file is FI Update

					var casetaskDT = new GlideRecord('x_baot_debt_sett_0_debt_task');
					casetaskDT.addQuery('number', sheet_case_task_id);
					casetaskDT.query();
					if (casetaskDT.next()) {

						casetaskDT.u_bulk_upload = true;
						target_record = sheet_case_task_id;

						if (casetaskDT.state == 7) { // If the state is Cancelled
							// add_log(source.sys_import_set, row_count, target_record, "Error", "ไม่สามารถแก้ไขใบงานที่สถานะใบงานเป็น Cancelled");
							add_log(source.sys_import_set, row_count, target_record, "Error", "ไม่สามารถแก้ไขใบงานได้ เนื่องจากสถานะใบงานถูกยกเลิก (Cancelled)");
							error_log = error_log + 1;
							continue;
						} else if (casetaskDT.state == 6 || casetaskDT.state == 3) { // If the state is Resolved or Closed
							if (isNotMember(userSysId_uploadby, casetaskDT.assignment_group)) {
								// add_log(source.sys_import_set, row_count, target_record, "Error", "ไม่มีสิทธินำเข้าข้อมูลสำหรับผู้ให้บริการที่เลือก");
								add_log(source.sys_import_set, row_count, target_record, "Error", "ผู้นำเข้าข้อมูลไม่อยู่ในกลุ่มผู้รับผิดชอบ");
								error_log = error_log + 1;
								continue;
							} else {
								if (check_date_valid(sheet_drd_report)) {
									casetaskDT.u_drd = set_date_object(sheet_drd_report);
									gs.info("[Excel Bulk Upload DET] วันที่ทำสัญญา " + set_date_object(sheet_drd_report));
								} else {
									add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "วันที่ทำสัญญา" ไม่ถูกต้อง โปรดระบุใน format DD-MM-YYYY');
									error_log = error_log + 1;
								}

								if (sheet_rdt_report == "" || check_date_valid(sheet_rdt_report)) {
									casetaskDT.u_rdt = set_date_object(sheet_rdt_report);
								} else {
									add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "รายงาน RDT" ไม่ถูกต้อง โปรดระบุใน format DD-MM-YYYY');
									error_log = error_log + 1;
								}

							}

							// Update Section
							if (error_log == 0) {
								casetaskDT.update();
								add_log(source.sys_import_set, row_count, target_record, "Update", "อัพเดทข้อมูลสำเร็จ (เฉพาะข้อมูล DRD และ/หรือ วันที่ทำสัญญา)");
							} else {
								error_log_all = 1;
								add_log(source.sys_import_set, row_count, target_record, "Error", "อัพเดตข้อมูลไม่สำเร็จ (Closed Case) มีข้อผิดพลาด " + String(error_log) + " ข้อ");
							}
						} else {
							add_log(source.sys_import_set, row_count, target_record, "Error", "ไม่สามารถแก้ไขใบงานที่สถานะใบงานไม่เป็น Resolved หรือ Closed");
							error_log = error_log + 1;
						}


					} else {
						error_log_all = 1;
						add_log(source.sys_import_set, row_count, target_record, "Error", "อัพเดตข้อมูลไม่สำเร็จ ไม่พบหมายเลข Case Task");
					}

				} else if (sheet_case_id != "" && sheet_case_task_id != "" && sheet_case_ref == "") {
					var caseGr = new GlideRecord('x_baot_debt_sett_0_case');
					caseGr.addQuery('number', sheet_case_id);

					caseGr.query();
					if (caseGr.next()) {
						var casetaskGr = new GlideRecord('x_baot_debt_sett_0_ect_debt_settlement');
						casetaskGr.addQuery('number', sheet_case_task_id);
						casetaskGr.query();
						if (casetaskGr.next()) {
							target_record = sheet_case_task_id;
							debt_project = casetaskGr.u_debt_project.getDisplayValue();

							if (casetaskGr.state == 7) {
								add_log(source.sys_import_set, row_count, target_record, "Error", "ไม่สามารถแก้ไขใบงานได้ เนื่องจากสถานะใบงานถูกยกเลิก (Cancelled)");
								error_log = error_log + 1;
								continue;
							}
							if (isNotMember(userSysId_uploadby, casetaskGr.assignment_group)) {
								add_log(source.sys_import_set, row_count, "", "Error", "ผู้นำเข้าข้อมูลไม่อยู่ในกลุ่มผู้รับผิดชอบ");
								error_log = error_log + 1;
								continue;
							} else if (casetaskGr.state == 3 || casetaskGr.state == 6) {
								// 3 == Closed && 6 == Resolved && 7 == Cancelled
								// รายงาน RDT
								if (sheet_rdt_report == "" || check_date_valid(sheet_rdt_report)) {
									if (check_date_valid(sheet_rdt_report)){
										casetaskGr.u_rdt = set_date_object(sheet_rdt_report);
									}
								} else {
									add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "รายงาน RDT" ไม่ถูกต้อง โปรดระบุใน format DD-MM-YYYY');
									error_log = error_log + 1;
								}

								// รายงาน DRD
								if (sheet_drd_report == "" || check_date_valid(sheet_drd_report)) {
									if (check_date_valid(sheet_drd_report)){
										casetaskGr.u_drd = set_date_object(sheet_drd_report);
									}
								} else {
									add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "วันที่ทำสัญญา" ไม่ถูกต้อง โปรดระบุใน format DD-MM-YYYY');
									error_log = error_log + 1;
								}


								// Update Section
								if (error_log == 0) {
									casetaskGr.update();
									add_log(source.sys_import_set, row_count, target_record, "Update", "อัพเดทข้อมูลสำเร็จ (เฉพาะข้อมูล DRD และ/หรือ วันที่ทำสัญญา)");
								} else {

									error_log_all = 1;

									add_log(source.sys_import_set, row_count, target_record, "Error", "อัพเดตข้อมูลไม่สำเร็จ (Close case) มีข้อผิดพลาด " + String(error_log) + " ข้อ");
								}
								continue;
							} else {


								// ผู้รับมอบหมาย
								casetaskGr.u_bulk_upload = false;
								if (sheet_assignee == "" && casetaskGr.state == 1 && sheet_state == "New") {
									add_log(source.sys_import_set, row_count, target_record, "Error", 'โปรดระบุข้อมูล "อีเมล"');
									error_log = error_log + 1;
								} else if (sheet_assignee == "" && casetaskGr.state != 1 && sheet_state != "New") {
									add_log(source.sys_import_set, row_count, target_record, "Error", 'โปรดระบุข้อมูล "อีเมล."');
									error_log = error_log + 1;
								} else if (check_user_valid_active('sys_user', 'email', sheet_assignee)) {
									var assigned_to = get_user_valid_active('sys_user', 'email', sheet_assignee);
									if (assigned_to == "Error") {
										add_log(source.sys_import_set, row_count, target_record, "Error", 'ไม่พบข้อมูลผู้ใช้อีเมลบนระบบ');
										error_log = error_log + 1;
									} else if (isNotMember(assigned_to, casetaskGr.assignment_group)) {
										add_log(source.sys_import_set, row_count, target_record, "Error", "อีเมลผู้รับมอบหมายไม่อยู่ในกลุ่มผู้รับผิดชอบ");
										error_log = error_log + 1;
									} else {
										casetaskGr.assigned_to = assigned_to;
										casetaskGr.u_bulk_upload = true;
									}

								} else {
									add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "ผู้รับมอบหมาย" ไม่ถูกต้อง');
									error_log = error_log + 1;
								}


								// ผลิตภัณฑ์
								if (check_select_valid_activePipe('x_baot_debt_sett_0_product_to_fair', "u_m2m_product_to_application.u_productSTARTSWITH" + sheet_product + "^x_baot_debt_sett_0_debt_fair.short_descriptionLIKE" + debt_project)) {
									var product = get_select_valid_active_app('u_m2m_product_to_application', 'u_product.u_product', sheet_product);
									if (product == "Error") {
										add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "ผลิตภัณฑ์" ไม่ถูกต้อง');
										error_log = error_log + 1;
									} else {
										casetaskGr.u_product = product;
									}

								} else {
									add_log(source.sys_import_set, row_count, target_record, "Error", 'โปรดระบุ "ผลิตภัณฑ์" ให้ถูกต้องตามโครงการ');
									error_log = error_log + 1;

								}

								//สถานะบัญชี
								if (sheet_account_status == "" && sheet_state == "Cancelled") {
									///

								} else if (check_select_valid_activePipe('x_baot_debt_sett_0_debt_to_fair', "u_m2m_debt_status_to_application.u_debt_statusSTARTSWITH" + sheet_account_status + "^x_baot_debt_sett_0_debt_fair.short_description=" + debt_project)) {
									var state_debt = get_select_valid_active_app('u_m2m_debt_status_to_application', 'u_debt_status.u_name', sheet_account_status);
									if (state_debt == "Error") {
										add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "สถานะบัญชี" ไม่ถูกต้อง');
										error_log = error_log + 1;
									} else {
										casetaskGr.u_state_debt = state_debt;
									}
								} else {
									add_log(source.sys_import_set, row_count, target_record, "Error", 'โปรดระบุ "สถานะบัญชี" ให้ถูกต้องตามโครงการ');
									error_log = error_log + 1;

								}

								// เลขที่ทะเบียนรถ validation rule
								// car
								if (carProduct.includes(sheet_product)) {
									gs.info(sheet_vehicle_number == "" && debt_project == "ทางด่วนแก้หนี้");
									if (sheet_vehicle_number == "" && debt_project == "ทางด่วนแก้หนี้") {
										add_log(source.sys_import_set, row_count, target_record, "Error", 'โปรดระบุ "เลขที่ทะเบียนรถ"');
										error_log = error_log + 1;
									} else if (sheet_vehicle_number != "") {
										casetaskGr.u_number_car = sheet_vehicle_number;
									} else {
										// โครงการคุณสู้เราช่วย ไม่ต้องใส่ก็ได้
									}
								}

								// จังหวัดที่จดทะเบียนรถ validation rule
								// car
								if (carProduct.includes(sheet_product)) {
									if (sheet_vehicle_number_province == "" && debt_project == "ทางด่วนแก้หนี้") {
										add_log(source.sys_import_set, row_count, target_record, "Error", 'โปรดระบุ "จังหวัดที่จดทะเบียนรถ"');
										error_log = error_log + 1;
									}
									// ใส่ค่าได้ทั้ง ทางด่วนแก้หนี้ และ โครงการคุณสู้ เราช่วย
									else if (check_user_valid_active('u_car_registration_province', 'u_vehicle', sheet_vehicle_number_province)) {
										var province_car = get_user_valid_active('u_car_registration_province', 'u_vehicle', sheet_vehicle_number_province);
										if (province_car == "Error") {
											add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "จังหวัดที่จดทะเบียนรถ" ไม่ถูกต้อง');
											error_log = error_log + 1;
										} else {
											casetaskGr.u_province_car = province_car;
										}
									}
									else {
										// โครงการคุณสู้ เราช่วย ไม่ต้องใส่ก็ได้
									}
								}

								// เลขที่บัตร/ เลขที่สัญญา
								// ใบงานเป็น เป็น inprogress หรือ Resolved จะต้องมีเลขที่บัตร/ เลขที่สัญญา ส่วนถ้าเป็น cancelled หรือ closed หรือ open ไม่ได้ required
								if (sheet_contract_number != "") {
									casetaskGr.u_number_contact = sheet_contract_number;
								} else if (sheet_contract_number == "" && (sheet_state == "Cancelled" || sheet_state == "Closed" || sheet_state == "Open" || sheet_state == "New")) {
									//do nothing
								} else {
									add_log(source.sys_import_set, row_count, target_record, "Error", 'โปรดระบุ "เลขที่บัตร/ เลขที่สัญญา"');
									error_log = error_log + 1;
								}

								// วันที่เริ่มติดต่อลูกหนี้ validation rule
								if (check_datetime_valid(sheet_date_of_start)) {
									casetaskGr.u_contact_debt = set_timezone_object(sheet_date_of_start);
								} else if (debt_project == "ทางด่วนแก้หนี้") {
									if (sheet_date_of_start == "" && (sheet_state == "New" || sheet_state == "Cancelled")) {
										//do nothing
									}
									else if (sheet_date_of_start == "") {
										add_log(source.sys_import_set, row_count, target_record, "Error", 'โปรดระบุ "วันที่เริ่มติดต่อลูกหนี้"');
										error_log = error_log + 1;
									}
									else {
										add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "วันที่เริ่มติดต่อลูกหนี้" ไม่ถูกต้อง โปรดระบุใน format DD-MM-YYYY HH:MM');
										error_log = error_log + 1;
									}

								} else if (debt_project == "โครงการคุณสู้ เราช่วย") {
									if (sheet_date_of_start == "") {
										//do nothing
									} else {
										add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "วันที่เริ่มติดต่อลูกหนี้" ไม่ถูกต้อง');
										error_log = error_log + 1;
									}

								}

								// Update state
								// Closed
								if (casetaskGr.state == 3) {
									// Closed = 3
									if (sheet_state != "Closed") {
										add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะใบงาน เนื่องจากใบงานปิดไปแล้ว");
										error_log = error_log + 1;
									}
								} else if (casetaskGr.state == 6) {
									if (sheet_state != "Resolved") {
										add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะใบงาน เนื่องจากใบงานเป็น Resolved");
										error_log = error_log + 1;
									}
								} else if (sheet_state == "Resolved") {
									if (sheet_assignee == "") {
										add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะใบงานไม่สามารถเป็น Resolved เนื่องจากผู้รับมอบหมายว่าง");
										error_log = error_log + 1;
									}
									if (sheet_result == '') {
										add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะใบงานไม่สามารถเป็น Resolved เนื่องจากผลการพิจารณาว่าง");
										error_log = error_log + 1;
									} else if ((sheet_product == '' || sheet_account_status == '' || sheet_contract_number == '' || sheet_total_debt == '' || sheet_principal == '') && (debt_project == "ทางด่วนแก้หนี้")) {
										// ผลิตภัณฑ์, สถานะบัญชี, เลขที่บัตร/ เลขที่สัญญา, ภาระหนี้รวม, เงินต้น
										add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะใบงานไม่สามารถเป็น Resolved เนื่องจากมี Mandatory Field ว่าง");
										error_log = error_log + 1;
									} else if ((sheet_product == '' || sheet_account_status == '' || sheet_contract_number == '' || sheet_principal == '') && (debt_project == "โครงการคุณสู้ เราช่วย")) {
										// ผลิตภัณฑ์, สถานะบัญชี, เลขที่บัตร/ เลขที่สัญญา, เงินต้น (บาท)
										add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะใบงานไม่สามารถเป็น Resolved เนื่องจากมี Mandatory Field ว่าง.");
										error_log = error_log + 1;
									} else if (check_select_state_active(sheet_state, sheet_result)) {

										var resolution_code = get_active_result(sheet_result);
										if (resolution_code == "Error") {
											add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "ผลการพิจารณา" ไม่ถูกต้อง');
											error_log = error_log + 1;
										} else {
											casetaskGr.u_resolution_code = resolution_code;
											casetaskGr.state = 6;
											casetaskGr.u_work_state = "เจ้าหน้าที่ดำเนินการเสร็จสิ้น";
											casetaskGr.u_bulk_upload = true;
											casetaskGr.close_notes = sheet_additional_info;

										}

									} else if (sheet_state == "") {
										//         //add code
									} else {
										add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "สถานะใบงาน" ไม่ถูกต้อง');
										error_log = error_log + 1;

									}

								} else if (sheet_state == "Cancelled") {
									if (sheet_result == '' || sheet_close == '') {
										add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะใบงานไม่สามารถเป็น Cancelled เนื่องจากผลการพิจารณาหรือ Close note ว่าง");
										error_log = error_log + 1;
									} else if (check_select_state_active(sheet_state, sheet_result)) {

										var resolution_code = get_active_result(sheet_result);
										if (resolution_code == "Error") {
											add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะใบงานไม่สามารถเป็น Cancelled เนื่องจากผลการผลการพิจารณาไม่ถูกต้อง");
											error_log = error_log + 1;
										} else {
											casetaskGr.u_resolution_code = resolution_code;
											casetaskGr.close_notes = sheet_close;
											casetaskGr.state = 7;
											casetaskGr.u_work_state = "ใบงานถูกยกเลิกโดยเจ้าหน้าที่";

											//closenote and other infoemation
											if (sheet_additional_info != "") {
												casetaskGr.close_notes += "\n" + sheet_additional_info;
											}
										}
									} else {
										add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะใบงานไม่สามารถเป็น Cancelled เนื่องจากผลการผลการพิจารณาไม่ถูกต้อง");
										error_log = error_log + 1;
									}
								} else if (sheet_state == "In Progress" || sheet_state == "Awaiting Info") {
									var intt = get_select_state_active(sheet_state);
									//In Progress
									if (intt == "2" || intt == 2) {
										if (sheet_assignee == '') {
											add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะใบงานไม่สามารถเป็น In Progress เนื่องจากผู้รับมอบหมายเป็นช่องว่าง");
											error_log = error_log + 1;
										} else {
											casetaskGr.state = 2;
											casetaskGr.u_work_state = "เจ้าหน้าที่กำลังดำเนินการ";
										}
									}
									//Awaiting Info
									else if (intt == "18" || intt == 18) {
										if (sheet_assignee == '') {
											add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะใบงานไม่สามารถเป็น Awaiting Info เนื่องจากผู้รับมอบหมายเป็นช่องว่าง");
											error_log = error_log + 1;
										} else {
											casetaskGr.state = 18;
											casetaskGr.u_work_state = "ขอข้อมูลเพิ่มเติม";
										}
									}

								} else {
									add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "สถานะใบงาน" ไม่ถูกต้อง');
									error_log = error_log + 1;
								}

								// ผลการพิจารณา
								if (sheet_result == "ได้ข้อสรุปกับลูกค้า") {
									casetaskGr.u_resolution_code = 100;
								} else if (sheet_result == "ไม่ได้ข้อสรุปกับลูกค้า") {
									casetaskGr.u_resolution_code = 200;
								} else if (sheet_result == "" && (sheet_state != "Closed")) {
									// do nothing
								} else if (sheet_state == "Cancelled" || sheet_state == "Resolved") {
									//do nothing
								} else {
									error_log = error_log + 1;
									add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "ผลการพิจารณา" ไม่ถูกต้อง');
								}

								// แนวทางการช่วยเหลือลูกหนี้
								if (sheet_result == "ได้ข้อสรุปกับลูกค้า") {
									if (check_select_valid_activePipe2('u_glideline_debtor', "u_name=" + sheet_guidelines_debtors, debt_project)) {
										var glideline_debt = get_select_valid_active_check_fair('u_glideline_debtor', 'u_name', sheet_guidelines_debtors, debt_project);
										if (glideline_debt == "Error") {
											add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "แนวทางการช่วยเหลือลูกหนี้" ไม่ถูกต้อง' + sheet_guidelines_debtors);
											error_log = error_log + 1;
										} else {
											casetaskGr.u_glideline_debt = glideline_debt;
										}
									} else if (sheet_guidelines_debtors == "") {
										add_log(source.sys_import_set, row_count, target_record, "Error", 'โปรดระบุ "แนวทางการช่วยเหลือลูกหนี้"' + sheet_guidelines_debtors);
										error_log = error_log + 1;
									} else {
										add_log(source.sys_import_set, row_count, target_record, "Error", 'โปรดระบุ "แนวทางการช่วยเหลือลูกหนี้" ให้ถูกต้องตามโครงการ' + sheet_guidelines_debtors);
										error_log = error_log + 1;
									}
								}

								// เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้
								if (sheet_result == "ไม่ได้ข้อสรุปกับลูกค้า") {
									if (check_select_valid_activePipe2('u_unable_help', "u_name=" + sheet_reason_not_help, debt_project)) {
										var reason_unable_help = get_select_valid_active_check_fair('u_unable_help', 'u_name', sheet_reason_not_help, debt_project);
										if (reason_unable_help == "Error") {
											add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้" ไม่ถูกต้อง');
											error_log = error_log + 1;
										} else {
											casetaskGr.u_reason_unable_help = reason_unable_help;
										}
									} else if (sheet_reason_not_help == "") {
										add_log(source.sys_import_set, row_count, target_record, "Error", 'โปรดระบุ "เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้"');
										error_log = error_log + 1;
									} else {
										add_log(source.sys_import_set, row_count, target_record, "Error", 'โปรดระบุ "เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้" ให้ถูกต้องตามโครงการ');
										error_log = error_log + 1;

									}
								}

								// ภาระหนี้รวม (บาท) validation rule
								if (check_money_valid(sheet_total_debt)) {
									casetaskGr.u_debt_burden = set_money_object(sheet_total_debt);
								} else {
									if (debt_project == "ทางด่วนแก้หนี้") {
										if (sheet_total_debt == "" && (sheet_state == "New" || sheet_state == "Cancelled")) {
											//do nothing
										} else if (sheet_total_debt == "") {
											add_log(source.sys_import_set, row_count, target_record, "Error", 'โปรดระบุ "ภาระหนี้รวม (บาท)"');
											error_log = error_log + 1;
										} else {
											add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "ภาระหนี้รวม (บาท)" ไม่ถูกต้อง');
											error_log = error_log + 1;
										}
									} else if (debt_project == "โครงการคุณสู้ เราช่วย") {
										if (sheet_total_debt == "") {
											//do nothing
										} else {
											add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "ภาระหนี้รวม (บาท)" ไม่ถูกต้อง');
											error_log = error_log + 1;
										}

									}
								}

								// เงินต้น (บาท)
								if (check_money_valid(sheet_principal)) {
									casetaskGr.u_principle = set_money_object(sheet_principal);
								} else if (sheet_principal == "" && (sheet_state == "New" || sheet_state == "Cancelled")) {
									/* empty */
								} else if (sheet_principal == "") {
									add_log(source.sys_import_set, row_count, target_record, "Error", 'โปรดระบุ "เงินต้น (บาท)"');
									error_log = error_log + 1;
								} else {
									add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "เงินต้น (บาท)" ไม่ถูกต้อง');
									error_log = error_log + 1;
								}

								// ภาระหนี้ที่ตกลงชำระ (บาท) validation rule
								if (sheet_result == "ได้ข้อสรุปกับลูกค้า") {
									// ใส่ค่าได้ทั้ง ทางด่วนแก้หนี้ และ โครงการคุณสู้ เราช่วย
									if (check_money_valid(sheet_debt_agreed)) {
										casetaskGr.u_debt_confirm = set_money_object(sheet_debt_agreed);
									}
									// log error ของ ทางด่วนแก้หนี้
									else if (debt_project == "ทางด่วนแก้หนี้") {
										if (sheet_debt_agreed == "" && (sheet_state == "New" || sheet_state == "Cancelled")) {
											//do nothing
										} else if (sheet_debt_agreed == "") {
											add_log(source.sys_import_set, row_count, target_record, "Error", 'โปรดระบุ "ภาระหนี้ที่ตกลงชำระ (บาท)"');
											error_log = error_log + 1;
										} else {
											add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "ภาระหนี้ที่ตกลงชำระ (บาท)" ไม่ถูกต้อง');
											error_log = error_log + 1;
										}
									}
									// log error ของโครงการคุณสู้ เราช่วย
									else if (debt_project == "โครงการคุณสู้ เราช่วย") {
										if (sheet_debt_agreed == "") {
											// do nothing
										} else {
											add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "ภาระหนี้ที่ตกลงชำระ (บาท)" ไม่ถูกต้อง');
											error_log = error_log + 1;
										}
									}
								}


								if (sheet_state == "Cancelled") {
									if (sheet_close == '') {
										add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะใบงานไม่สามารถเป็น Cancelled เนื่องจาก Close note ว่าง");
										error_log = error_log + 1;
									} else if (check_select_state_active(sheet_state, sheet_result)) {

										casetaskGr.close_notes = sheet_close;

									} else {
										add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะใบงานไม่สามารถเป็น Cancelled เนื่องจากผลการผลการพิจารณาไม่ถูกต้อง");
										error_log = error_log + 1;
									}
								}

								// จำนวนงวดที่ชำระ (เดือน) validation rule
								// ถ้าไม่ได้ข้อสรุปกับลูกค้า field นี้จะไม่ปรากฎ
								if (sheet_result == "ได้ข้อสรุปกับลูกค้า") {
									// ทางด่วนแก้หนี้ required จำนวนงวดที่ต้องชำระ
									if (debt_project == "ทางด่วนแก้หนี้") {
										if (sheet_no_payment == "" && (sheet_state == "New" || sheet_state == "Cancelled")) {
											//do nothing
										} else if (sheet_no_payment == "") {
											add_log(source.sys_import_set, row_count, target_record, "Error", 'โปรดระบุ "จำนวนงวดที่ชำระ (เดือน)"');
											error_log = error_log + 1;
										} else {
											var intValue = parseInt(sheet_no_payment, 10);
											if (isNaN(intValue) || intValue < 0) {
												add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "จำนวนงวดที่ชำระ (เดือน)" ไม่ถูกต้อง');
												error_log = error_log + 1;
											} else {
												casetaskGr.u_installment = parseInt(sheet_no_payment);
											}

										}
									}
									else if (debt_project == "โครงการคุณสู้ เราช่วย") {
										// โครงการคุณสู้เราช่วย ไม่เป็น required
										if (sheet_no_payment == "") {
											//do nothing
										} else {
											var intValue = parseInt(sheet_no_payment, 10);
											if (isNaN(intValue) || intValue < 0) {
												add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "จำนวนงวดที่ชำระ (เดือน)" ไม่ถูกต้อง');
												error_log = error_log + 1;
											} else {
												casetaskGr.u_installment = parseInt(sheet_no_payment);
											}
										}
									}
								}

								// ค่างวดต่อเดือน (บาท) validation rule
								// ปรากฎเมื่อ ผลการพิจารณา = ได้สรุปกับลูกค้า
								if (sheet_result == "ได้ข้อสรุปกับลูกค้า") {
									// ใส่ค่าได้ทั้งโครงการทางด่วนแก้หนี้ และ โครงการคุณสู้ เราช่วย
									if (check_money_valid(sheet_monthly)) {
										casetaskGr.u_amount_month = set_money_object(sheet_monthly);
									} 
									// log error
									else if (debt_project == "ทางด่วนแก้หนี้") {
										if (sheet_monthly == "" && (sheet_state == "New" || sheet_state == "Cancelled")) {
											//do nothing
										}
										else if (sheet_monthly == "") {
											add_log(source.sys_import_set, row_count, target_record, "Error", 'โปรดระบุ "ค่างวดต่อเดือน (บาท)"');
											error_log = error_log + 1;
										}
										else {
											add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "ค่างวดต่อเดือน (บาท)" ไม่ถูกต้อง');
											error_log = error_log + 1;
										}
									} 
									// log error
									else if (debt_project == "โครงการคุณสู้ เราช่วย") {
										if (sheet_monthly == "") {
											//do nothing
										} 
										else {
											add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "ค่างวดต่อเดือน (บาท)" ไม่ถูกต้อง');
											error_log = error_log + 1;
										}

									}
								}

								// รายงาน RDT
								if (sheet_rdt_report == "" || check_date_valid(sheet_rdt_report)) {
									if (check_date_valid(sheet_rdt_report)){
										casetaskGr.u_rdt = set_date_object(sheet_rdt_report);
									}
									
								} else {
									add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "รายงาน RDT" ไม่ถูกต้อง โปรดระบุใน format DD-MM-YYYY');
									error_log = error_log + 1;
								}

								// รายงาน DRD
								if (sheet_drd_report == "" || check_date_valid(sheet_drd_report)) {
									if (check_date_valid(sheet_drd_report)){
										casetaskGr.u_drd = set_date_object(sheet_drd_report);
									}
								} else {
									add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "วันที่ทำสัญญา" ไม่ถูกต้อง โปรดระบุใน format DD-MM-YYYY');
									error_log = error_log + 1;
								}

								//closenote and other infoemation
								if (sheet_state == "Cancelled") {
									if (sheet_additional_info != "") {
										casetaskGr.close_notes += "\n" + sheet_additional_info;
									} else if (sheet_additional_info == "") {
										casetaskGr.close_notes = sheet_close;
									}
								} else {
									casetaskGr.close_notes = sheet_additional_info;
								}


								// Update Section
								if (error_log == 0) {
									casetaskGr.update();
									gs.info('u_resolution_code' + casetaskGr.u_resolution_code);
									add_log(source.sys_import_set, row_count, target_record, "Update", "อัพเดทข้อมูลสำเร็จ");
									// test(targetGR, sheet_case_task_id);
								} else {

									error_log_all = 1;

									add_log(source.sys_import_set, row_count, target_record, "Error", "อัพเดทข้อมูลไม่สำเร็จ มีข้อผิดพลาด " + String(error_log) + " ข้อ");
								}
							}
						} else {
							add_log(source.sys_import_set, row_count, target_record, "Error", "ไม่เจอ Case Task ที่จะ Update : " + sheet_case_task_id);
							error_log = error_log + 1;
							error_log_all = 1;
							//ไม่เจอ Case task

						}

					} else {
						add_log(source.sys_import_set, row_count, target_record, "Error", "ไม่เจอ Case ID ที่จะ Update : " + sheet_case_id);
						error_log = error_log + 1;
						error_log_all = 1;
						//ไม่เจอ Case task

					}
				}
				// Insert Task----------------------------------------------------------------------------------------------------------------------------------------------------
				else if (sheet_case_id == "" && sheet_case_task_id == "" && sheet_case_ref != "") {

					var caseref;

					var casedebtGr = new GlideRecord('x_baot_debt_sett_0_ect_debt_settlement');
					if (casedebtGr.get(get_user_valid_active('x_baot_debt_sett_0_ect_debt_settlement', 'number', sheet_case_ref))) {
						caseref = casedebtGr.parent;
					}

					var caseGr = new GlideRecord('x_baot_debt_sett_0_case');
					caseGr.addQuery('sys_id', caseref);

					caseGr.setLimit(1);
					caseGr.query();
					if (caseGr.next()) {
						//case cancelled value is 7
						//case closed value is 3
						if (caseGr.state == "3") {
							add_log(source.sys_import_set, row_count, caseGr.number, "Error", "ไม่สามารถสร้างใบงานได้ เนื่องจากสถานะใบ ref case task งานถูกปิดแล้ว (Closed)");
							error_log = error_log + 1;
							continue;
						} else if (caseGr.state == "7") {
							add_log(source.sys_import_set, row_count, caseGr.number, "Error", "ไม่สามารถสร้างใบงานได้ เนื่องจากสถานะใบ ref case task งานเป็นยกเลิก (Cancelled)");
							error_log = error_log + 1;
							continue;
						}

						//insert field
						var casetaskGr = new GlideRecord('x_baot_debt_sett_0_ect_debt_settlement');
						casetaskGr.initialize();

						casetaskGr.u_bulk_upload = true;
						casetaskGr.number = sheet_case_task_id;

						target_record = sheet_case_relate;



						if (check_user_valid_active('x_baot_debt_sett_0_ect_debt_settlement', 'number', sheet_case_ref)) {
							// Case Detail (optional)
							var casedubGr = new GlideRecord('x_baot_debt_sett_0_ect_debt_settlement');
							if (casedubGr.get(get_user_valid_active('x_baot_debt_sett_0_ect_debt_settlement', 'number', sheet_case_ref))) {
								casetaskGr.u_debt_detail = casedubGr.u_debt_detail;
							}
						} else {
							add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "Case Detail" ไม่ถูกต้อง');
							error_log = error_log + 1;

						}


						var cGr = new GlideRecord('x_baot_debt_sett_0_ect_debt_settlement');
						//  gs.info('sheet_case_ref' + sheet_case_ref);
						cGr.addQuery('number', sheet_case_ref);

						cGr.query();
						if (cGr.next()) {
							sheet_title = cGr.short_description;
							// sheet_status = cGr.u_work_state;
							//duke edit
							sheet_sla_duedate = cGr.u_sla_due_date;
							sheet_duedate = cGr.due_date;
							sheet_phone_backup = cGr.u_secondary_phone;
							sheet_email_backup = cGr.u_secondary_email;
							sheet_responsible_group = cGr.assignment_group;
							sheet_guidelines_provider = cGr.u_offer_provider;
							sheet_provider = cGr.u_provider;
							sheet_project_debt_record = cGr.u_debt_project;
						};
						//จังหวัด
						casetaskGr.u_province = cGr.u_province;
						//ผู้ติดต่อ
						casetaskGr.u_exact_consumer = cGr.u_exact_consumer;
						//ผู้ขอรับบริการ
						casetaskGr.u_consumer = cGr.u_consumer;
						//ข้อมูลที่ต้องการแจ้ง ธปท (ถ้ามี)
						// casetaskGr.close_notes = cGr.close_notes;
						// หมายเลข Case

						casetaskGr.u_parent_case = caseGr.getUniqueValue();
						casetaskGr.u_parent = caseGr.getUniqueValue();
						casetaskGr.parent = caseGr.sys_id;
						casetaskGr.u_debt_project = sheet_project_debt_record;

						//show on ticket list page
						casetaskGr.u_show_on_ticket_list_page = true;

						// เรื่อง
						casetaskGr.short_description = sheet_title;

						// สถานะการดำเนินงาน
						casetaskGr.u_work_state = 'รอเจ้าหน้าที่เข้ามารับงาน';

						// SLA Due date
						casetaskGr.u_sla_due_date = sheet_sla_duedate;

						casetaskGr.assignment_group = sheet_responsible_group;

						casetaskGr.due_date = sheet_duedate;

						var project_debt_record = cGr.u_debt_project.getDisplayValue();
						debt_project = cGr.u_debt_project.getDisplayValue();

						if (isNotMember(userSysId_uploadby, sheet_responsible_group)) {
							add_log(source.sys_import_set, row_count, target_record, "Error", "ผู้นำเข้าข้อมูลไม่อยู่ในกลุ่มผู้รับผิดชอบ");
							error_log = error_log + 1;
							continue;
						}



						if (check_user_valid_active('sys_user', 'email', sheet_assignee)) {
							var assigned_to = get_user_valid_active('sys_user', 'email', sheet_assignee);
							if (assigned_to == "Error") {
								add_log(source.sys_import_set, row_count, target_record, "Error", 'ไม่พบข้อมูลผู้ใช้ Email บนระบบ');
								error_log = error_log + 1;
							} else {
								if (isNotMember(assigned_to, sheet_responsible_group)) {
									add_log(source.sys_import_set, row_count, target_record, "Error", 'ไม่พบข้อมูลผู้ใช้ Email บนระบบ');
									error_log = error_log + 1;
								} else {
									casetaskGr.assigned_to = assigned_to;
								}
							}
						} else if (sheet_assignee == "") {
							//add code
						} else {
							add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "อีเมล" ไม่เป็นไปตามรูปแบบ เช่น example@domain.xyz');
							error_log = error_log + 1;

						}

						casetaskGr.u_secondary_phone = sheet_phone_backup;

						casetaskGr.u_secondary_email = sheet_email_backup;

						casetaskGr.u_provider = sheet_provider;

						// ผลิตภัณฑ์
						if (check_select_valid_activePipe('x_baot_debt_sett_0_product_to_fair', "u_m2m_product_to_application.u_productSTARTSWITH" + sheet_product + "^x_baot_debt_sett_0_debt_fair.short_descriptionLIKE" + project_debt_record)) {
							var product = get_select_valid_active_app('u_m2m_product_to_application', 'u_product.u_product', sheet_product);
							if (product == "Error") {
								add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "ผลิตภัณฑ์" ไม่ถูกต้อง');
								error_log = error_log + 1;
							} else {
								casetaskGr.u_product = product;
							}
						} else if (sheet_product == "") {
							//add code
						} else {
							add_log(source.sys_import_set, row_count, target_record, "Error", 'โปรดระบุ "ผลิตภัณฑ์" ให้ถูกต้องตามโครงการ');
							error_log = error_log + 1;

						}

						//สถานะบัญชี
						if (check_select_valid_activePipe('x_baot_debt_sett_0_debt_to_fair', "u_m2m_debt_status_to_application.u_debt_statusSTARTSWITH" + sheet_account_status + "^x_baot_debt_sett_0_debt_fair.short_description=" + project_debt_record)) {
							var state_debt = get_select_valid_active_app('u_m2m_debt_status_to_application', 'u_debt_status.u_name', sheet_account_status);
							if (state_debt == "Error") {
								add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "สถานะบัญชี" ไม่ถูกต้อง');
								error_log = error_log + 1;
							} else {
								casetaskGr.u_state_debt = state_debt;
							}
						} else if (sheet_account_status == "") { } else {
							add_log(source.sys_import_set, row_count, target_record, "Error", 'โปรดระบุ "สถานะบัญชี" ให้ถูกต้องตามโครงการ');
							error_log = error_log + 1;

						}


						// เลขที่ทะเบียนรถ
						// car
						if (carProduct.includes(sheet_product)) {
							casetaskGr.u_number_car = sheet_vehicle_number;
						}

						// จังหวัดที่จดทะเบียนรถ
						// car
						if (carProduct.includes(sheet_product)) {
							if (check_user_valid_active('u_car_registration_province', 'u_vehicle', sheet_vehicle_number_province)) {
								var province_car = get_user_valid_active('u_car_registration_province', 'u_vehicle', sheet_vehicle_number_province);
								if (province_car == "Error") {
									add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "จังหวัดที่จดทะเบียนรถ" ไม่ถูกต้อง');
									error_log = error_log + 1;
								} else {
									casetaskGr.u_province_car = province_car;
								}
							} else if (sheet_vehicle_number_province == "") {
								/* empty */
							} else {
								add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "จังหวัดที่จดทะเบียนรถ" ไม่ถูกต้อง');
								error_log = error_log + 1;
							}
						}

						// แนวทางที่ต้องการเสนอให้ผู้บริการทางการเงินพิจารณา
						// ถูกดึงมาจากใบงานแม่
						casetaskGr.u_offer_provider = sheet_guidelines_provider;

						// เหตุผลหรือคำอธิบายประกอบคำขอ
						casetaskGr.u_provider_verified = true;


						// Case ที่เกี่ยวข้อง
						if (check_user_valid_active('task', 'number', sheet_case_relate_2)) {
							var major_case = get_user_valid_active('task', 'number', sheet_case_relate_2);
							if (major_case == "Error") {
								add_log(source.sys_import_set, row_count, target_record, "Error", 'ไม่พบข้อมูล Case ที่เกี่ยวข้อง บนระบบ' + sheet_case_relate_2);
								error_log = error_log + 1;
							} else {
								casetaskGr.u_major_case = major_case;
							}
						} else if (sheet_case_relate_2 == "") {
							//add code
						} else {
							add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "Case ที่เกี่ยวข้อง" ไม่ถูกต้อง' + sheet_case_relate_2);
							error_log = error_log + 1;

						}

						// เลขที่บัตร/ เลขที่สัญญา
						// ใบงานเป็น เป็น inprogress หรือ Resolved จะต้องมีเลขที่บัตร/ เลขที่สัญญา ส่วนถ้าเป็น cancelled หรือ closed หรือ open ไม่ได้ required
						if (sheet_contract_number != "") {
							casetaskGr.u_number_contact = sheet_contract_number;
						} 
						// else if (sheet_contract_number == "" && (sheet_state == "Cancelled" || sheet_state == "Closed" || sheet_state == "Open" || sheet_state == "New")) {
						// 	//do nothing
						// }
						else {
							add_log(source.sys_import_set, row_count, target_record, "Error", 'โปรดระบุ "เลขที่บัตร/ เลขที่สัญญา"');
							error_log = error_log + 1;
						}

						// วันที่เริ่มติดต่อลูกหนี้ validation rule
						if (check_datetime_valid(sheet_date_of_start)) {
							casetaskGr.u_contact_debt = set_timezone_object(sheet_date_of_start);
						} else if (debt_project == "ทางด่วนแก้หนี้") {
							if (sheet_date_of_start == "" && (sheet_state == "New" || sheet_state == "Cancelled")) {
								//do nothing
							}
							else if (sheet_date_of_start == "") {
								add_log(source.sys_import_set, row_count, target_record, "Error", 'โปรดระบุ "วันที่เริ่มติดต่อลูกหนี้"');
								error_log = error_log + 1;
							}
							else {
								add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "วันที่เริ่มติดต่อลูกหนี้" ไม่ถูกต้อง โปรดระบุใน format DD-MM-YYYY HH:MM');
								error_log = error_log + 1;
							}

						} else if (debt_project == "โครงการคุณสู้ เราช่วย") {
							if (sheet_date_of_start == "") {
								//do nothing
							} else {
								add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "วันที่เริ่มติดต่อลูกหนี้" ไม่ถูกต้อง');
								error_log = error_log + 1;
							}

						}


						if (sheet_state == "In Progress" || sheet_state == "Awaiting Info") {
							// ผลการพิจารณาสามารเป็นเป็น  ["ได้ข้อสรุปกับลูกค้า","ไม่ได้ข้อสรุปกับลูกค้า", null];
							// ผลการพิจารณา
							if (sheet_result == "ได้ข้อสรุปกับลูกค้า") {
								casetaskGr.u_resolution_code = 100;
							} else if (sheet_result == "ไม่ได้ข้อสรุปกับลูกค้า") {
								casetaskGr.u_resolution_code = 200;
							} else if (sheet_result != "") {
								error_log = error_log + 1;
								add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "ผลการพิจารณา" ไม่ถูกต้อง');
							}

							var intt = get_select_state_active(sheet_state);
							// value inProgrss:2, awaitingInfo:18
							if (intt == "2" || intt == 2) {
								if (sheet_assignee == '') {
									add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะใบงานไม่สามารถเป็น In Progress เนื่องจากผู้รับมอบหมายเป็นค่าว่าง");
									error_log = error_log + 1;
								} else {
									casetaskGr.state = 2;
									casetaskGr.u_work_state = "เจ้าหน้าที่กำลังดำเนินการ";
								}
								// เขียนเพิ่มจากเดิมไม่มี awaiting info
							} else if (intt == "18" || intt == 18) {
								casetaskGr.state = 18;
								casetaskGr.u_work_state = "รอข้อมูลเพิ่มเติม";
							}

						} else if (sheet_state == "New" || sheet_state == "Open") {

							add_log(source.sys_import_set, row_count, target_record, "Error", "สถานะใบงานไม่สามารถเป็น New หรือ Open ");
							error_log = error_log + 1;

						} else if (sheet_state == "Resolved") {
							// ผลการพิจารณาสามารเป็นเป็น  ["ได้ข้อสรุปกับลูกค้า","ไม่ได้ข้อสรุปกับลูกค้า"];
							if (sheet_result == '') {
								add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะใบงานไม่สามารถเป็น Resolved เนื่องจากผลการพิจารณาว่าง");
								error_log = error_log + 1;
							} else if ((sheet_product == '' || sheet_account_status == '' || sheet_contract_number == '' || sheet_total_debt == '' || sheet_principal == '') && (debt_project == "ทางด่วนแก้หนี้")) {
								// ผลิตภัณฑ์, สถานะบัญชี, เลขที่สัญญา, ภาระหนี้รวม, เงินต้น
								add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะใบงานไม่สามารถเป็น Resolved เนื่องจากมี Mandatory Field ว่าง");
								error_log = error_log + 1;
							} else if ((sheet_product == '' || sheet_account_status == '' || sheet_contract_number == '' || sheet_principal == '') && (debt_project == "โครงการคุณสู้ เราช่วย")) {
								// ผลิตภัณฑ์, สถานะบัญชี, เลขที่สัญญา, เงินต้น
								add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะใบงานไม่สามารถเป็น Resolved เนื่องจากมี Mandatory Field ว่าง.");
								error_log = error_log + 1;
							} else if (check_select_state_active(sheet_state, sheet_result)) {
								var resolution_code = get_active_result(sheet_result);
								if (resolution_code == "Error") {
									add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "ผลการพิจารณา" ไม่ถูกต้อง');
									error_log = error_log + 1;
								} else {
									casetaskGr.u_resolution_code = resolution_code;
									casetaskGr.state = 6;
									casetaskGr.u_work_state = "เจ้าหน้าที่ดำเนินการเสร็จสิ้น";
									casetaskGr.u_bulk_upload = true;
									casetaskGr.close_notes = sheet_additional_info;
								}
							} else {
								add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "สถานะใบงาน" ไม่ถูกต้อง');
								error_log = error_log + 1;
							}
						} else if (sheet_state == "Cancelled") {
							// ผลการพิจารณาสามารเป็นเป็น  ["ยกเลิกโดยเจ้าหหน้าที่", "ยอดหนี้เท่ากับศูนย์", "ลูกค้ายกเลิกคำขอ", "คำขอซ้ำ", "ปิดบัญชีแล้ว", "บัญชีถูกขายไปแล้ว", "ไม่พบบัญชีลูกค้า"];
							if (sheet_result == '' || sheet_close == '') {
								add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะใบงานไม่สามารถเป็น Cancelled เนื่องจากผลการพิจารณาหรือ Close note ว่าง");
								error_log = error_log + 1;
							} else if (check_select_state_active(sheet_state, sheet_result)) {

								var resolution_code = get_active_result(sheet_result);
								if (resolution_code == "Error") {
									add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะใบงานไม่สามารถเป็น Cancelled เนื่องจากผลการผลการพิจารณาไม่ถูกต้อง");
									error_log = error_log + 1;
								} else {
									casetaskGr.u_resolution_code = resolution_code;
									casetaskGr.close_notes = sheet_close;
									casetaskGr.state = 7;
									casetaskGr.u_work_state = "ใบงานถูกยกเลิกโดยเจ้าหน้าที่";

									// เพิ่ม close note กับข้อมูลเพิ่มเติม append เพิ่ม
									if (sheet_additional_info != "") {
										casetaskGr.close_notes += "\n" + sheet_additional_info;
									}

								}
							} else {
								add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะใบงานไม่สามารถเป็น Cancelled เนื่องจากผลการผลการพิจารณาไม่ถูกต้อง");
								error_log = error_log + 1;
							}
						} else if (sheet_state == "Closed") {
							error_log = error_log + 1;
							add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะใบงานไม่สามารถเป็น Closed ได้จากการ Duplicate ใบงาน");
						} else {
							error_log = error_log + 1;
							add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "สถานะใบงาน" ไม่ถูกต้อง');
						}



						// แนวทางการช่วยเหลือลูกหนี้
						if (sheet_result == "ได้ข้อสรุปกับลูกค้า") {
							if (check_select_valid_activePipe2('u_glideline_debtor', "u_name=" + sheet_guidelines_debtors, debt_project)) {
								var glideline_debt = get_select_valid_active_check_fair('u_glideline_debtor', 'u_name', sheet_guidelines_debtors, debt_project);
								if (glideline_debt == "Error") {
									add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "แนวทางการช่วยเหลือลูกหนี้" ไม่ถูกต้อง' + sheet_guidelines_debtors);
									error_log = error_log + 1;
								} else {
									casetaskGr.u_glideline_debt = glideline_debt;
								}
							} else if (sheet_guidelines_debtors == "") {
								add_log(source.sys_import_set, row_count, target_record, "Error", 'โปรดระบุ "แนวทางการช่วยเหลือลูกหนี้"' + sheet_guidelines_debtors);
								error_log = error_log + 1;
							} else {
								add_log(source.sys_import_set, row_count, target_record, "Error", 'โปรดระบุ "แนวทางการช่วยเหลือลูกหนี้" ให้ถูกต้องตามโครงการ' + sheet_guidelines_debtors);
								error_log = error_log + 1;
							}
						}


						// เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้
						if (sheet_result == "ไม่ได้ข้อสรุปกับลูกค้า") {
							if (check_select_valid_activePipe2('u_unable_help', "u_name=" + sheet_reason_not_help, debt_project)) {
								var reason_unable_help = get_select_valid_active_check_fair('u_unable_help', 'u_name', sheet_reason_not_help, debt_project);
								if (reason_unable_help == "Error") {
									add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้" ไม่ถูกต้อง');
									error_log = error_log + 1;
								} else {
									casetaskGr.u_reason_unable_help = reason_unable_help;
								}
							} else if (sheet_reason_not_help == "") {
								add_log(source.sys_import_set, row_count, target_record, "Error", 'โปรดระบุ "เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้"');
								error_log = error_log + 1;
							} else {
								add_log(source.sys_import_set, row_count, target_record, "Error", 'โปรดระบุ "เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้" ให้ถูกต้องตามโครงการ');
								error_log = error_log + 1;

							}
						}
						
						// ภาระหนี้รวม (บาท) validation rule
						if (check_money_valid(sheet_total_debt)) {
							casetaskGr.u_debt_burden = set_money_object(sheet_total_debt);
						} else {
							if (debt_project == "ทางด่วนแก้หนี้") {
								if (sheet_total_debt == "" && (sheet_state == "New" || sheet_state == "Cancelled")) {
									//do nothing
								} else if (sheet_total_debt == "") {
									add_log(source.sys_import_set, row_count, target_record, "Error", 'โปรดระบุ "ภาระหนี้รวม (บาท)"');
									error_log = error_log + 1;
								} else {
									add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "ภาระหนี้รวม (บาท)" ไม่ถูกต้อง');
									error_log = error_log + 1;
									//do nothing
								}
							} else if (debt_project == "โครงการคุณสู้ เราช่วย") {
								if (sheet_total_debt == "") {
									//do nothing
								} else {
									add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "ภาระหนี้รวม (บาท)" ไม่ถูกต้อง');
									error_log = error_log + 1;
								}

							}
						}

						// เงินต้น (บาท)
						if (check_money_valid(sheet_principal)) {
							casetaskGr.u_principle = set_money_object(sheet_principal);
						} else if (sheet_principal == "" && (sheet_state == "New" || sheet_state == "Cancelled")) {
							/* empty */
						} else if (sheet_state == "Cancelled") {
							/* empty */
						} else if (sheet_principal == "") {
							add_log(source.sys_import_set, row_count, target_record, "Error", 'โปรดระบุ "เงินต้น (บาท)"');
							error_log = error_log + 1;
						} else {
							add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "เงินต้น (บาท)" ไม่ถูกต้อง');
							error_log = error_log + 1;
						}

						// ภาระหนี้ที่ตกลงชำระ (บาท) validation rule
						if (sheet_result == "ได้ข้อสรุปกับลูกค้า") {
							// ใส่ค่าได้ทั้ง ทางด่วนแก้หนี้ และ โครงการคุณสู้ เราช่วย
							if (check_money_valid(sheet_debt_agreed)) {
								casetaskGr.u_debt_confirm = set_money_object(sheet_debt_agreed);
							}
							// log error ของ ทางด่วนแก้หนี้
							else if (debt_project == "ทางด่วนแก้หนี้") {
								if (sheet_debt_agreed == "" && (sheet_state == "New" || sheet_state == "Cancelled")) {
									//do nothing
								} else if (sheet_debt_agreed == "") {
									add_log(source.sys_import_set, row_count, target_record, "Error", 'โปรดระบุ "ภาระหนี้ที่ตกลงชำระ (บาท)"');
									error_log = error_log + 1;
								} else {
									add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "ภาระหนี้ที่ตกลงชำระ (บาท)" ไม่ถูกต้อง');
									error_log = error_log + 1;
								}
							}
							// log error ของโครงการคุณสู้ เราช่วย
							else if (debt_project == "โครงการคุณสู้ เราช่วย") {
								if (sheet_debt_agreed == "") {
									// do nothing
								} else {
									add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "ภาระหนี้ที่ตกลงชำระ (บาท)" ไม่ถูกต้อง');
									error_log = error_log + 1;
								}
							}
						}

						// จำนวนงวดที่ชำระ (เดือน) validation rule
						// ถ้าไม่ได้ข้อสรุปกับลูกค้า field นี้จะไม่ปรากฎ
						if (sheet_result == "ได้ข้อสรุปกับลูกค้า") {
							// ทางด่วนแก้หนี้ required จำนวนงวดที่ต้องชำระ
							if (debt_project == "ทางด่วนแก้หนี้") {
								if (sheet_no_payment == "" && (sheet_state == "New" || sheet_state == "Cancelled")) {
									//do nothing
								} else if (sheet_no_payment == "") {
									add_log(source.sys_import_set, row_count, target_record, "Error", 'โปรดระบุ "จำนวนงวดที่ชำระ (เดือน)"');
									error_log = error_log + 1;
								} else {
									var intValue = parseInt(sheet_no_payment, 10);
									if (isNaN(intValue) || intValue < 0) {
										add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "จำนวนงวดที่ชำระ (เดือน)" ไม่ถูกต้อง');
										error_log = error_log + 1;
									} else {
										casetaskGr.u_installment = parseInt(sheet_no_payment);
									}

								}
							}
							else if (debt_project == "โครงการคุณสู้ เราช่วย") {
								// โครงการคุณสู้เราช่วย ไม่เป็น required
								if (sheet_no_payment == "") {
									//do nothing
								} else {
									var intValue = parseInt(sheet_no_payment, 10);
									if (isNaN(intValue) || intValue < 0) {
										add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "จำนวนงวดที่ชำระ (เดือน)" ไม่ถูกต้อง');
										error_log = error_log + 1;
									} else {
										casetaskGr.u_installment = parseInt(sheet_no_payment);
									}
								}
							}
						}

						// ค่างวดต่อเดือน (บาท) validation rule
						// ปรากฎเมื่อ ผลการพิจารณา = ได้สรุปกับลูกค้า
						if (sheet_result == "ได้ข้อสรุปกับลูกค้า") {
							// ใส่ค่าได้ทั้งโครงการทางด่วนแก้หนี้ และ โครงการคุณสู้ เราช่วย
							if (check_money_valid(sheet_monthly)) {
								casetaskGr.u_amount_month = set_money_object(sheet_monthly);
							} 
							// log error
							else if (debt_project == "ทางด่วนแก้หนี้") {
								if (sheet_monthly == "" && (sheet_state == "New" || sheet_state == "Cancelled")) {
									//do nothing
								}
								else if (sheet_monthly == "") {
									add_log(source.sys_import_set, row_count, target_record, "Error", 'โปรดระบถ "ค่างวดต่อเดือน (บาท)"');
									error_log = error_log + 1;
								}
								else {
									add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "ค่างวดต่อเดือน (บาท)" ไม่ถูกต้อง');
									error_log = error_log + 1;
								}
							} 
							// log error
							else if (debt_project == "โครงการคุณสู้ เราช่วย") {
								if (sheet_monthly == "") {
									//do nothing
								} 
								else {
									add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "ค่างวดต่อเดือน (บาท)" ไม่ถูกต้อง');
									error_log = error_log + 1;
								}

							}
						}

						// รายงาน RDT
						if (sheet_rdt_report == "" || check_date_valid(sheet_rdt_report)) {
							if (check_date_valid(sheet_rdt_report)){
								casetaskGr.u_rdt = set_date_object(sheet_rdt_report);
							}
						} else {
							add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "รายงาน RDT" ไม่ถูกต้อง โปรดระบุใน format DD-MM-YYYY');
							error_log = error_log + 1;
						}


						// รายงาน DRD
						if (sheet_drd_report == "" || check_date_valid(sheet_drd_report)) {
							if (check_date_valid(sheet_drd_report)){
								casetaskGr.u_drd = set_date_object(sheet_drd_report);
							}
						} else {
							add_log(source.sys_import_set, row_count, target_record, "Error", 'รูปแบบข้อมูล "วันที่ทำสัญญา" ไม่ถูกต้อง โปรดระบุใน format DD-MM-YYYY');
							error_log = error_log + 1;
						}


						// ข้อมูลที่ต้องการแจ้ง ธปท. เพิ่มเติม (ถ้ามี)
						casetaskGr.u_other_informantion = sheet_additional_info;

						if (error_log == 0) {

							var grd = new GlideRecord('x_baot_debt_sett_0_ect_debt_settlement');
							grd.addQuery('number', targetGR.u_ref_case_task);
							grd.query();

							if (grd.next()) {

								var parent = grd.parent.toString();
								var parentNumber = grd.parent.number.toString();
							}

							gs.info('parent' + parent);
							gs.info('parentNumber' + parentNumber);


							var gr = new GlideRecord('x_baot_debt_sett_0_ect_debt_settlement');
							gr.addQuery('parent', parent);
							gr.query();
							var rowCountNumber = gr.getRowCount() + countNumber;


							var type = 'casetask';
							var recordNumber = new global.MF_ChildRecordNumberPadding().padNumber(rowCountNumber, parentNumber, type);
							gs.info('recordNumber' + recordNumber);

							casetaskGr.number = recordNumber;
							casetaskGr.insert();

							var case_number = recordNumber;
							add_log(source.sys_import_set, row_count, case_number, "Insert", "เพิ่มข้อมูลสำเร็จ");
						} else {
							error_log_all = 1;
							add_log(source.sys_import_set, row_count, target_record, "Error", "เพิ่มข้อมูลไม่สำเร็จ มีข้อผิดพลาด " + String(error_log) + " ข้อ");

						}

					} else {
						//ไม่เจอ Case
						// add_log(source.sys_import_set, row_count, target_record, "Error", "ค้นหา Case ID ไม่เจอ! " + sheet_case_relate);
						add_log(source.sys_import_set, row_count, target_record, "Error", "ไม่พบข้อมูล Case Task ตามหมายเลขที่ระบุเป็นเป็น Ref Case Task : " + sheet_case_relate);
						error_log = error_log + 1;
						error_log_all = 1;

					}
				} else {
					add_log(source.sys_import_set, row_count, target_record, "Error", "ผิด Format การอัพเดตข้อมูล ");
					error_log = error_log + 1;
					error_log_all = 1;
					continue;
				}
			}
		}

		var log_total = 0;
		var log_error_total = 0;
		var selectGrRecord = new GlideRecord("x_baot_debt_sett_0_import_excel_logs");
		selectGrRecord.addQuery("importsetref", source.sys_import_set);
		selectGrRecord.query();
		if (selectGrRecord.next()) {
			log_total = selectGrRecord.getRowCount();
		}
		gs.info("[Excel Bulk Upload]source.sys_import_set : " + source.sys_import_set);
		var erorrGrRecord = new GlideRecord("x_baot_debt_sett_0_import_excel_logs");
		erorrGrRecord.addQuery("importsetref", source.sys_import_set);
		erorrGrRecord.addQuery("state", 'ERROR');
		erorrGrRecord.query();
		if (erorrGrRecord.next()) {
			log_error_total = erorrGrRecord.getRowCount();
		}

		var row_table_importattachment = new GlideRecord('x_baot_debt_sett_0_import_excel_attachments');
		row_table_importattachment.addQuery("sys_id", importExcelAttachmentSysId);
		row_table_importattachment.query();
		if (row_table_importattachment.next()) {


			row_table_importattachment.setValue("importset_link", source.sys_import_set);
			// gs.info("row_table_importattachment : " + row_table_importattachment);
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
		gs.info("[Excel Bulk Upload Transform] Complete ");
	} catch (error) {

		gs.info("[Excel Bulk Upload Error catch] error: " + error);


		// Update Import Set in Table Import Attachment
		var sysid_table_importattachment = importExcelAttachmentSysId;
		var row_table_importattachment = new GlideRecord('x_baot_debt_sett_0_import_excel_attachments');
		row_table_importattachment.addQuery("sys_id", sysid_table_importattachment);
		row_table_importattachment.query();
		if (row_table_importattachment.next()) {

			//gs.info("[Excel Bulk Upload] Import Set add(sys_import_set): " + import_set.sys_id + " | " + row_importset.sys_id);
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
			gs.info("[Excel Bulk Upload Error catch] Record row_table_importattachment not found!");
		}


	}

})(source, map, log, target);

function set_timezone_object(datetime_string_input) {

	var datePart = datetime_string_input.split(" ")[0];
	var timePart = datetime_string_input.split(" ")[1];
	var dateArray = datePart.split("-");

	var formattedDate = dateArray[2] + "-" + dateArray[1] + "-" + dateArray[0] + " " + timePart + ":00";

	var gdt1 = new GlideDateTime(formattedDate);
	gdt1.addSeconds(-25200);


	return gdt1;
}

function check_datetime_valid(date_input) {


	var regex = /^(\d{2})-(\d{2})-(\d{4}) (\d{2}):(\d{2})$/;
	//  
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

function set_date_object(date_string_input) {
	var regex = /^(\d{2})-(\d{2})-(\d{4})$/;
	var match = date_string_input.match(regex);
	var day = parseInt(match[1], 10);
	var month = parseInt(match[2], 10);
	var year = parseInt(match[3], 10);

	var dayWithZero = String(day).padStart(2, '0');
	var monthWithZero = String(month).padStart(2, '0');

	return year + "-" + month + "-" + day;
}

function check_date_valid(date_input) {
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

	var day = parseInt(match2[1], 10);
	var month = parseInt(match2[2], 10);
	var year = parseInt(match2[3], 10);

	if (day < 1 || day > 31 || month < 1 || month > 12) {
		return false;
	}

	var daysInMonth = [31, (year % 4 === 0 && year % 100 !== 0) || (year % 400 === 0) ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
	if (day > daysInMonth[month - 1]) {
		return false;
	}

	return true;
}

function check_idcard_valid(id_number) {
	var regex = /^[0-9]{13}$/;
	return regex.test(id_number);
}

function set_money_object(money_input) {
	var x = money_input.replace("฿", "");
	var y = x.replace(",", "");
	return "THB;" + y;
}

function check_money_valid(amount) {
	// var regex = /^(฿)?(\d{1,3}(,\d{3})*|\d+)(\.\d{1,2})?$/;
	var regex = /^(฿)?(\d{1,3}(,\d{3})*|\d+)(\.\d+)?$/;
	return regex.test(amount);
}

function check_email_valid(email) {
	var regex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
	return regex.test(email);
}

function check_phone_valid(phoneNumber) {
	gs.info('phoneNumber' + phoneNumber);
	var regex = /^66\d{9}$/;
	var text = String(phoneNumber).replace(" ", "");
	gs.info('text' + text);
	return regex.test(text);

}

function check_number_valid(inputNumber) {
	var regex = /^\d+$/;
	return regex.test(inputNumber);
}

function check_select_valid_active(table_select, choice_col, choice_select) {
	//gs.info('Select > ' + table_select + ' ' + choice_col + ' ' + choice_select);
	var selectGr = new GlideRecord(table_select);
	selectGr.addQuery(choice_col, choice_select);
	selectGr.addQuery('u_active', true);
	selectGr.setLimit(1);
	selectGr.query();
	if (selectGr.next()) {
		return true;
	} else {
		return false;
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

function check_active(table_select, choice_select) {
	//gs.info('Select > ' + table_select + ' ' + choice_col + ' ' + choice_select);
	var selectGr = new GlideRecord(table_select);
	selectGr.addQuery('u_name', choice_select);
	selectGr.addQuery('u_active', true);
	selectGr.setLimit(1);
	selectGr.query();
	if (selectGr.next()) {
		return true;
	} else {
		return false;
	}
}



function check_select_state_active(choice_select, sheet_result) {

	if (choice_select == "Resolved") {
		var selectGr = new GlideRecord('sys_choice');
		selectGr.addQuery('name', 'u_debt_settlement_master_case_task');
		selectGr.addQuery('label', sheet_result);
		selectGr.addQuery('element', 'u_resolution_code');
		selectGr.setLimit(1);
		selectGr.query();
		if (selectGr.next()) {
			if (selectGr.value == '100' || selectGr.value == '200') {
				return true;
			} else {
				return false;
			}
		} else {
			return false;
		}
	} else {
		var selectGr1 = new GlideRecord('sys_choice');
		selectGr1.addQuery('name', 'u_debt_settlement_master_case_task');
		selectGr1.addQuery('label', sheet_result);
		selectGr1.addQuery('element', 'u_resolution_code');
		selectGr1.setLimit(1);
		selectGr1.query();
		if (selectGr1.next()) {
			if (selectGr1.value == '10' || selectGr1.value == '500' || selectGr1.value == '9' || selectGr1.value == '8' || selectGr1.value == '800' || selectGr1.value == '700' || selectGr1.value == '600')
				return true;
		} else {
			return false;
		}
	}


}

function check_select_states_active(choice_select) {

	var selectGr = new GlideRecord('sys_choice');
	selectGr.addQuery('name', 'u_master_case_task');
	selectGr.addQuery('label', choice_select);
	selectGr.addQuery('element', 'state');
	selectGr.setLimit(1);
	selectGr.query();
	if (selectGr.next()) {

		return true;
	} else {
		return false;
	}



}

function get_select_state_active(choice_select) {
	var selectGr = new GlideRecord('sys_choice');
	selectGr.addQuery('name', 'u_master_case_task');
	selectGr.addQuery('label', choice_select);
	selectGr.addQuery('element', 'state');

	selectGr.setLimit(1);
	selectGr.query();
	if (selectGr.next()) {

		return selectGr.value;

	} else {
		return false;;
	}
}

function get_select_valid_active(table_select, choice_col, choice_select) {
	var selectGr = new GlideRecord(table_select);
	selectGr.addQuery(choice_col, choice_select);
	selectGr.addEncodedQuery("u_application_management=adbb01c71bbb65d080bb844ee54bcbeb");
	selectGr.addQuery("u_active", "true");
	selectGr.setLimit(1);
	selectGr.query();
	if (selectGr.next()) {
		return selectGr.getUniqueValue();
	} else {
		return "Error";
	}
}

function get_select_valid_active_app(table_select, choice_col, choice_select) {
	var selectGr = new GlideRecord(table_select);
	selectGr.addQuery(choice_col, choice_select);
	selectGr.addQuery('u_application_management', 'adbb01c71bbb65d080bb844ee54bcbeb');
	selectGr.addQuery("u_active", "true");
	selectGr.setLimit(1);
	selectGr.query();
	if (selectGr.next()) {
		return selectGr.getUniqueValue();
	} else {
		return "Error";
	}
}

//not use
function get_active(choice_select) {
	var selectGr = new GlideRecord('u_glideline_debtor');
	selectGr.addQuery('u_name', choice_select);
	selectGr.addQuery('u_active', true);
	selectGr.setLimit(1);
	selectGr.query();
	if (selectGr.next()) {

		return selectGr.sys_id;

	} else {
		return "Error";
	}
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

function check_user_valid_active(table_select, choice_col, choice_select) {
	//gs.info('Select > ' + table_select + ' ' + choice_col + ' ' + choice_select);
	var selectGr = new GlideRecord(table_select);
	selectGr.addQuery(choice_col, choice_select);
	// selectGr.addQuery('u_active', 'True');
	selectGr.setLimit(1);
	selectGr.query();
	if (selectGr.next()) {
		return true;
	} else {
		return false;
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
	//logrow.upload_by = userSysId;
	logrow.insert();
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

function formatDate(inputDate) {
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
		formattedDate = "Error";
	}
	return formattedDate;
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
	gs.info("[Excel Bulk Upload IsNotMember]user: " + userSysId + " group: " + groupSysId);
	return memberOfGroup;
}

function sleep(ms) {
	var unixtime_ms = new Date().getTime();
	while (new Date().getTime() < unixtime_ms + ms) { }

}