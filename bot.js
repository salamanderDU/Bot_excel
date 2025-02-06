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
		var carProduct = ["สินเชื่อเช่าซื้อรถยนต์ / Top-up / รถแลกเงินแบบโอนเล่ม", "สินเชื่อเช่าซื้อรถจักรยานยนต์ / Top-up / รถแลกเงินโอนเล่ม", "จำนำทะเบียนรถยนต์ (ไม่โอนเล่ม)", "จำนำทะเบียนรถจักรยานยนต์ (ไม่โอนเล่ม)","สินเชื่อเช่าซื้อรถยนต์", "สินเชื่อเช่าซื้อรถจักรยานยนต์", "สินเชื่อจำนำทะเบียนรถยนต์", "สินเชื่อจำนำทะเบียนรถจักรยานยนต์"];


		var grSA = new GlideRecord('sys_attachment');
		grSA.addEncodedQuery("table_sys_idSTARTSWITHd4093f2747d3c250ab9eb5c8736d43da");
		grSA.setLimit(1);
		grSA.orderByDesc("sys_created_on");
		grSA.query();
		if (grSA.next()) {
		
		// var gr1 = new GlideRecord('sys_attachment');
		// gr1.addQuery('table_name', 'x_baot_debt_sett_0_import_excel_attachments');
		// gr1.orderByDesc('sys_created_on');
		// gr1.query();
		// if (gr1.next()) {
			
			var sysid_table_importattachment1 = grSA.getValue("file_name").split("-")[0];

			gs.info("sysid_table_importattachment1: "+grSA.getValue("file_name").split("-")[0]);

			var row_table_importattachment1 = new GlideRecord('x_baot_debt_sett_0_import_excel_attachments');
			row_table_importattachment1.addQuery('sys_id', sysid_table_importattachment1);
			row_table_importattachment1.query();
			if (row_table_importattachment1.next()) {
				userSysId_uploadby = row_table_importattachment1.upload_by;
				importExcelAttachmentSysId = row_table_importattachment1.sys_id;
			}
		// }

		}

		gs.info("userSysId_uploadby: "+userSysId_uploadby);
		gs.info("importExcelAttachmentSysId: "+importExcelAttachmentSysId);

		// gs.info("duke"+JSON.stringify(source));
		// //new get userSysId_uploadby
		// var grIM = new GlideRecord("x_baot_debt_sett_0_import_excel_attachments");
		// importExcel = source.sys_import_set;
		// gs.info("[Excel Bulk Upload importExcel]importExcel :"+importExcel);
		// grIM.addQuery("importset_link", importExcel);
		// grIM.query();
		// // gs.info(JSON.stringify(grIM));
		// if(grIM.next()){
		// 	gs.info("userSysId_uploadby = "+grIM.upload_by);
		// 	userSysId_uploadby = grIM.upload_by;
		// }else{
		// 	gs.info("[Excel Bulk Uplaod] not in");
		// }


		gs.info("[Excel Bulk Uplaod]get user: "+userSysId_uploadby);

		var recordsLimit = gs.getProperty('x_baot_debt_sett_0.excelRecordsLimit');
		var countCondition = false;
		if (recordsLimit < 1) countCondition = count < 1;
		else countCondition = count < 1 || count > recordsLimit;

		if (countCondition) {
			var gr = new GlideRecord('sys_attachment');
			gr.addQuery('table_name', 'x_baot_debt_sett_0_import_excel_attachments');
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


			//error_log = error_log + 1;
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
				gs.info("V3Log_เรื่อง : " + sheet_title + " product : " + sheet_product);

				var sheet_account_status = targetGR.u_สถานะบ_ญช_;
				if (sheet_account_status == '' || sheet_account_status == null) {
					//    gs.info('sheet_guidelines_debtorssssssssssssllllllllllds');

				} else {
					sheet_account_status = targetGR.u_สถานะบ_ญช_.split("|")[0].trim();

				}
				//  gs.info('sheet_guidelines_debtorssssssssseeeeeeeeeeeeeeeedlds');

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
				if (sheet_reason_not_help == '' || sheet_guidelines_debtors == null) {
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
				var sheet_rtd_report = targetGR.u_รายงาน_rdt;
				// sheet_rtd_report = formatDate(sheet_rtd_report);
				// รายงาน DRD
				var sheet_drd_report = targetGR.u_รายงาน_drd ? targetGR.u_รายงาน_drd : targetGR.u_ว_นท__ทำส_ญญา;
				// sheet_drd_report = formatDate(sheet_drd_report);
				// ข้อมูลที่ต้องการแจ้ง ธปท. เพิ่มเติม (ถ้ามี)
				var sheet_additional_info = targetGR.u_ข_อม_ลท__ต_อ__มเต_ม__ถ_าม__;

				var sheet_state = targetGR.u_สถานะใบงาน;
				var sheet_cancel_state = targetGR.u_ประเภทการแก_ไข;
				var sheet_close = targetGR.u_close_note;



				if (sheet_case_id != "" && sheet_case_task_id != "" && sheet_case_ref == "") {
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
								add_log(source.sys_import_set, row_count, target_record, "Error", "ใบงานเป็น Cancelled ไม่สามารถแก้ไขได้");
								continue;
							}
							if (isNotMember(userSysId_uploadby, casetaskGr.assignment_group)) {
								add_log(source.sys_import_set, row_count, "", "Error", "ผิดพลาด ผู้นำเข้าข้อมูลไม่อยู่ในกลุ่มผู้รับผิดชอบ");
								error_log = error_log + 1;
								continue;
							} else if (casetaskGr.state == 3 || casetaskGr.state == 6) {
								// 3 == Closed && 6 == Resolved && 7 == Cancelled
								// รายงาน RDT

								if (sheet_rtd_report) {
									sheet_rtd_report = formatDate(sheet_rtd_report);
									if (check_date_valid(sheet_rtd_report)) {
										casetaskGr.u_rdt = set_date_object(sheet_rtd_report);
									} else {
										add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด รายงาน RDT");
										error_log = error_log + 1;
									}
								}


								// รายงาน DRD
								if (sheet_drd_report) {
									sheet_drd_report = formatDate(sheet_drd_report);
									if (check_date_valid(sheet_drd_report)) {
										casetaskGr.u_drd = set_date_object(sheet_drd_report);
									} else {
										add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด วันที่ทำสัญญา");
										error_log = error_log + 1;
									}
								}

								// Update Section
								if (error_log == 0) {
									casetaskGr.update();
									add_log(source.sys_import_set, row_count, target_record, "Update", "อัพเดทข้อมูลสำเร็จ (เฉพาะข้อมูล DRD และ/หรือ วันที่ทำสัญญา)");
								} else {

									error_log_all = 1;

									add_log(source.sys_import_set, row_count, target_record, "Error", "อัพเดตข้อมูลไม่สำเร็จ (Close Case) มีข้อผิดพลาด " + String(error_log) + " ข้อ");
								}
								continue;
							} else {


								// ผู้รับมอบหมาย
								casetaskGr.u_bulk_upload = false;
								if (sheet_assignee == "" && casetaskGr.state == 1 && sheet_state == "New") {
									//add code
								} else if (sheet_assignee == "" && casetaskGr.state != 1 && sheet_state != "New") {


									add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด ผู้รับมอบหมาย");
									error_log = error_log + 1;


								} else if (check_user_valid_active('sys_user', 'email', sheet_assignee)) {
										var assigned_to = get_user_valid_active('sys_user', 'email', sheet_assignee);
										if (assigned_to == "Error") {
											add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด ผู้รับมอบหมาย");
											error_log = error_log + 1;
										} else if (isNotMember(assigned_to, casetaskGr.assignment_group)) {
												add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด ผู้รับมอบหมาย");
												error_log = error_log + 1;
											} else {
												casetaskGr.assigned_to = assigned_to;
												casetaskGr.u_bulk_upload = true;
											}

									} else {
										add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด ผู้รับมอบหมาย");
										error_log = error_log + 1;
									}


								// ผลิตภัณฑ์
								if (check_select_valid_activePipe('x_baot_debt_sett_0_product_to_fair', "u_m2m_product_to_application.u_productSTARTSWITH" + sheet_product + "^x_baot_debt_sett_0_debt_fair.short_descriptionLIKE" + debt_project)) {
									var product = get_select_valid_active('u_m2m_product_to_application', 'u_product.u_product', sheet_product);
									if (product == "Error") {
										add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด ผลิตภัณฑ์");
										error_log = error_log + 1;
									} else {
										casetaskGr.u_product = product;
									}

								} else {
									add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด ผลิตภัณฑ์");
									error_log = error_log + 1;

								}

								//สถานะบัญชี
								if (sheet_account_status == "" && sheet_state == "Cancelled") {
									///

								} else if (check_select_valid_activePipe('x_baot_debt_sett_0_debt_to_fair', "u_m2m_debt_status_to_application.u_debt_statusSTARTSWITH" + sheet_account_status + "^x_baot_debt_sett_0_debt_fair.short_description=" + debt_project)) {
									var state_debt = get_select_valid_active('u_m2m_debt_status_to_application', 'u_debt_status.u_name', sheet_account_status);
									if (state_debt == "Error") {
										add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะบัญชี");
										error_log = error_log + 1;
									} else {
										casetaskGr.u_state_debt = state_debt;
									}
								} else {
									add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะบัญชี");
									error_log = error_log + 1;

								}

								// เลขที่ทะเบียนรถ validation rule
								// car
								if (carProduct.includes(sheet_product)) {
									gs.info(sheet_vehicle_number == "" && debt_project == "ทางด่วนแก้หนี้");
									if (sheet_vehicle_number == "" && debt_project == "ทางด่วนแก้หนี้") {
										add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด เลขที่ทะเบียนรถ");
										error_log = error_log + 1;
									} else if (sheet_vehicle_number != "") {
										casetaskGr.u_number_car = sheet_vehicle_number;
									} else {
										// add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด เลขที่ทะเบียนรถ");
										// error_log = error_log + 1;
									}
								}

								// จังหวัดที่จดทะเบียนรถ validation rule
								// car
								if (carProduct.includes(sheet_product)) {
									if (sheet_vehicle_number_province == "" && debt_project == "ทางด่วนแก้หนี้") {
										add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด จังหวัดที่จดทะเบียนรถ");
										error_log = error_log + 1;
									} else if (check_user_valid_active('u_car_registration_province', 'u_vehicle', sheet_vehicle_number_province)) {
										var province_car = get_user_valid_active('u_car_registration_province', 'u_vehicle', sheet_vehicle_number_province);
										if (province_car == "Error") {
											add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด จังหวัดที่จดทะเบียนรถ");
											error_log = error_log + 1;
										} else {
											casetaskGr.u_province_car = province_car;
										}
									} else {
										// add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด จังหวัดที่จดทะเบียนรถ");
										// error_log = error_log + 1;

									}
								}

								// เลขที่บัตร/ เลขที่สัญญา
								if (sheet_contract_number != "") {
									casetaskGr.u_number_contact = sheet_contract_number;
								} else if (sheet_contract_number == "" && (sheet_state == "New" || sheet_state == "Cancelled")) {
										// add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด เลขที่บัตร/ เลขที่สัญญา");
										// error_log = error_log + 1;
										//add code
									} else {
										add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด เลขที่บัตร/ เลขที่สัญญา");
										error_log = error_log + 1;

									}
								//fix
								// วันที่เริ่มติดต่อลูกหนี้ validation rule
								if (check_datetime_valid(sheet_date_of_start)) {
									casetaskGr.u_contact_debt = set_timezone_object(sheet_date_of_start);
								} else if (debt_project == "ทางด่วนแก้หนี้") {
										if (sheet_date_of_start == "" && (sheet_state == "New" || sheet_state == "Cancelled")) {
											//do nothing
										} else {
											add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด วันที่เริ่มติดต่อลูกหนี้");
											error_log = error_log + 1;
										}

									} else if (debt_project == "โครงการคุณสู้ เราช่วย") {
										if (sheet_date_of_start == "") {
											//do nothing
										} else {
											add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด วันที่เริ่มติดต่อลูกหนี้");
											error_log = error_log + 1;
										}

									}
								
								// Update state
								// Closed
								// jump
								if (casetaskGr.state == 3) {
									// Closed = 3
									if (sheet_state != "Closed") {
										add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะใบงาน เนื่องจากใบงานปิดไปแล้ว");
										error_log = error_log + 1;
									}
								} else if (casetaskGr.state == 6) {
									if (sheet_state != "Resolved") {
										add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะใบงาน เนื่องจากใบงานปิดไปแล้ว");
										error_log = error_log + 1;
									}
								}  else if (sheet_state == "Resolved") {
										if(sheet_assignee == ""	){
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
											add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะใบงานไม่สามารถเป็น Resolved เนื่องจากมี Mandatory Field ว่าง");
											error_log = error_log + 1;
										} else if (check_select_state_active(sheet_state, sheet_result)) {

												var resolution_code = get_active_result(sheet_result);
												if (resolution_code == "Error") {
													add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะใบงาน");
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
													add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะใบงาน");
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
									add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะใบงาน");
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
								}else {
									error_log = error_log + 1;
									add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด ผลการพิจารณา");
								}

								// แนวทางการช่วยเหลือลูกหนี้
								if (sheet_result == "ได้ข้อสรุปกับลูกค้า") {
									if (check_select_valid_activePipe2('u_glideline_debtor', "u_name=" + sheet_guidelines_debtors, debt_project)) {
										var glideline_debt = get_select_valid_active_check_fair('u_glideline_debtor', 'u_name', sheet_guidelines_debtors, debt_project);
										if (glideline_debt == "Error") {
											add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด แนวทางการช่วยเหลือลูกหนี้ " + sheet_guidelines_debtors);
											error_log = error_log + 1;
										} else {
											casetaskGr.u_glideline_debt = glideline_debt;
										}
									} else if (sheet_guidelines_debtors == "") {
											add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด แนวทางการช่วยเหลือลูกหนี้ " + sheet_guidelines_debtors);
											error_log = error_log + 1;
										} else {
											add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด แนวทางการช่วยเหลือลูกหนี้ " + sheet_guidelines_debtors);
											error_log = error_log + 1;
										}
								}

								// เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้
								if (sheet_result == "ไม่ได้ข้อสรุปกับลูกค้า") {
									if (check_select_valid_activePipe2('u_unable_help', "u_name=" + sheet_reason_not_help, debt_project)) {
										var reason_unable_help = get_select_valid_active_check_fair('u_unable_help', 'u_name', sheet_reason_not_help, debt_project);
										if (reason_unable_help == "Error") {
											add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้");
											error_log = error_log + 1;
										} else {
											casetaskGr.u_reason_unable_help = reason_unable_help;
										}
									} else if (sheet_reason_not_help == "") {
											add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้");
											error_log = error_log + 1;
										} else {
											add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้");
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
										} else {
											add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด ภาระหนี้รวม (บาท)");
											error_log = error_log + 1;
											//do nothing
										}
									} else if (debt_project == "โครงการคุณสู้ เราช่วย") {
										if (sheet_total_debt == "") {
											//do nothing
										} else {
											add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด ภาระหนี้รวม (บาท)");
											error_log = error_log + 1;
										}

									}
								}

								// เงินต้น (บาท)
								if (check_money_valid(sheet_principal)) {
									casetaskGr.u_principle = set_money_object(sheet_principal);
								} else if (sheet_principal == "" && (sheet_state == "New" || sheet_state == "Cancelled")) { /* empty */ } else if (sheet_state == "Cancelled") { /* empty */ } else {
										add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด เงินต้น (บาท)");
										error_log = error_log + 1;

									}

								/// ภาระหนี้ที่ตกลงชำระ (บาท) validation rule
								if (check_money_valid(sheet_debt_agreed)) {
									casetaskGr.u_debt_confirm = set_money_object(sheet_debt_agreed);
								} else if (debt_project == "ทางด่วนแก้หนี้") {
										if (sheet_debt_agreed == "" && (sheet_state == "New" || sheet_state == "Cancelled")) {
											//do nothing
										} else {
											add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด ภาระหนี้ที่ตกลงชำระ (บาท)");
											error_log = error_log + 1;
										}
									} else if (debt_project == "โครงการคุณสู้ เราช่วย") {
										if (sheet_debt_agreed == "") {
											// do nothing
										} else {
											add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด ภาระหนี้ที่ตกลงชำระ (บาท)");
											error_log = error_log + 1;
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
								if (sheet_result == "ไม่ได้ข้อสรุปกับลูกค้า") {
									///
								} else if (debt_project == "ทางด่วนแก้หนี้") {
										if (sheet_no_payment == "" && (sheet_state == "New" || sheet_state == "Cancelled")) {
											//do nothing
										} else if (sheet_no_payment == "") {
											add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด จำนวนงวดที่ชำระ (เดือน)");
											error_log = error_log + 1;
										} else {
											var intValue = parseInt(sheet_no_payment, 10);
											if (isNaN(intValue) || intValue < 0) {
												add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด จำนวนงวดที่ชำระ (เดือน)");
												error_log = error_log + 1;
											} else {
												casetaskGr.u_installment = parseInt(sheet_no_payment);
											}

										}
									} else if (debt_project == "โครงการคุณสู้ เราช่วย") {
										//just insert
										if (sheet_no_payment == "") {
											//do nothing
										} else {
											var intValue = parseInt(sheet_no_payment, 10);
											if (isNaN(intValue) || intValue < 0) {
												add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด จำนวนงวดที่ชำระ (เดือน)");
												error_log = error_log + 1;
											} else {
												casetaskGr.u_installment = parseInt(sheet_no_payment);
											}
										}
									}

								// ค่างวดต่อเดือน (บาท) validation rule
								if (check_money_valid(sheet_monthly)) {
									casetaskGr.u_amount_month = set_money_object(sheet_monthly);
								} else {
									if (debt_project == "ทางด่วนแก้หนี้") {
										if (sheet_monthly == "" && (sheet_state == "New" || sheet_state == "Cancelled")) {
											//do nothing
										} else {
											add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด ค่างวดต่อเดือน (บาท)");
											error_log = error_log + 1;
										}
									} else if (debt_project == "โครงการคุณสู้ เราช่วย") {
										if (sheet_monthly == "") {
											//do nothing
										} else {
											add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด ค่างวดต่อเดือน (บาท)");
											error_log = error_log + 1;
										}

									}
								}

								// รายงาน RDT
								if (sheet_rtd_report) {
									sheet_rtd_report = formatDate(sheet_rtd_report);
									if (check_date_valid(sheet_rtd_report)) {
										gs.info("[Excel Bulk Upload time] rdt report if " + sheet_rtd_report);
										casetaskGr.u_rdt = set_date_object(sheet_rtd_report);
									} else {
										add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด รายงาน RDT");
										error_log = error_log + 1;
									}
								}
								// รายงาน DRD

								if (sheet_drd_report) {
									sheet_drd_report = formatDate(sheet_drd_report);
									if (check_date_valid(sheet_drd_report)) {
										gs.info("[Excel Bulk Upload time] drd report if" + sheet_drd_report);
										casetaskGr.u_drd = set_date_object(sheet_drd_report);
									} else {
										add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด วันที่ทำสัญญา");
										error_log = error_log + 1;
									}
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
							gs.info("duke: " + sheet_case_task_id);
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
				else if (sheet_case_ref != "") {

					var caseref;

					var casedebtGr = new GlideRecord('x_baot_debt_sett_0_ect_debt_settlement');
					//casedebtGr.addQuery('sys_id', caseref);
					if (casedebtGr.get(get_user_valid_active('x_baot_debt_sett_0_ect_debt_settlement', 'number', sheet_case_ref))) {
						caseref = casedebtGr.parent;
					}


					var caseGr = new GlideRecord('x_baot_debt_sett_0_case');
					caseGr.addQuery('sys_id', caseref);

					caseGr.setLimit(1);
					caseGr.query();
					if (caseGr.next()) {
						if (caseGr.state == 3) {
							add_log(source.sys_import_set, row_count, caseGr.number, "Error", "ผิดพลาด สถานะใบงานปิด");
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
							add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด Case Detail");
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

						if(isNotMember(userSysId_uploadby,sheet_responsible_group)){
							add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด ผู้นำเข้าข้อมูลไม่อยู่ในกลุ่มผู้รับผิดชอบ");
							error_log = error_log + 1;
							continue;
						}



						if (check_user_valid_active('sys_user', 'email', sheet_assignee)) {
							var assigned_to = get_user_valid_active('sys_user', 'email', sheet_assignee);
							if (assigned_to == "Error") {
								add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด ผู้รับมอบหมาย");
								error_log = error_log + 1;
							} else {
								if (isNotMember(assigned_to, sheet_responsible_group)) {
									add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด ผู้รับมอบหมาย");
									error_log = error_log + 1;
								} else {
									casetaskGr.assigned_to = assigned_to;
								}
							}
						} else if (sheet_assignee == "") {
								//add code
							} else {
								add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด ผู้รับมอบหมาย");
								error_log = error_log + 1;

							}

						casetaskGr.u_secondary_phone = sheet_phone_backup;

						casetaskGr.u_secondary_email = sheet_email_backup;

						casetaskGr.u_provider = sheet_provider;

						// ผลิตภัณฑ์
						if (check_select_valid_activePipe('x_baot_debt_sett_0_product_to_fair', "u_m2m_product_to_application.u_productSTARTSWITH" + sheet_product + "^x_baot_debt_sett_0_debt_fair.short_descriptionLIKE" + project_debt_record)) {
							var product = get_select_valid_active_app('u_m2m_product_to_application', 'u_product.u_product', sheet_product);
							if (product == "Error") {
								add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด ผลิตภัณฑ์");
								error_log = error_log + 1;
							} else {
								casetaskGr.u_product = product;
							}
						} else if (sheet_product == "") {
								//add code
							} else {
								add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด ผลิตภัณฑ์");
								error_log = error_log + 1;

							}

						//สถานะบัญชี
						if (check_select_valid_activePipe('x_baot_debt_sett_0_debt_to_fair', "u_m2m_debt_status_to_application.u_debt_statusSTARTSWITH" + sheet_account_status + "^x_baot_debt_sett_0_debt_fair.short_description=" + project_debt_record)) {
							var state_debt = get_select_valid_active_app('u_m2m_debt_status_to_application', 'u_debt_status.u_name', sheet_account_status);
							if (state_debt == "Error") {
								add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะบัญชี");
								error_log = error_log + 1;
							} else {
								casetaskGr.u_state_debt = state_debt;
							}
						} else if (sheet_account_status == "") { } else {
								add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะบัญชี");
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
									add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด จังหวัดที่จดทะเบียนรถ");
									error_log = error_log + 1;
								} else {
									casetaskGr.u_province_car = province_car;
								}
							} else if (sheet_vehicle_number_province == "") { /* empty */ } else {
									add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด จังหวัดที่จดทะเบียนรถ");
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
								add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด Case ที่เกี่ยวข้อง " + sheet_case_relate_2);
								error_log = error_log + 1;
							} else {
								casetaskGr.u_major_case = major_case;
							}
						} else if (sheet_case_relate_2 == "") {
								//add code
							} else {
								add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด Case ที่เกี่ยวข้อง " + sheet_case_relate_2);
								error_log = error_log + 1;

							}

						// Ref Case Task
						// 

						// เลขที่บัตร/ เลขที่สัญญา
						if (sheet_contract_number != "") {
							casetaskGr.u_number_contact = sheet_contract_number;
						} else if (sheet_contract_number == "") {
								//add code
							} else {
								add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด เลขที่บัตร/ เลขที่สัญญา");
								error_log = error_log + 1;

							}

						// วันที่เริ่มติดต่อลูกหนี้
						if (check_datetime_valid(sheet_date_of_start)) {
							casetaskGr.u_contact_debt = set_timezone_object(sheet_date_of_start);
						} else if (sheet_date_of_start == "") {
								//add code
						} else {
							add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด วันที่เริ่มติดต่อลูกหนี้");
							error_log = error_log + 1;

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
								add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด ผลการพิจารณา");
							}

                            var intt = get_select_state_active(sheet_state);
							// value inProgrss:2, awaitingInfo:18
                            if (intt == "2" || intt == 2) {
                                if ( sheet_assignee == '') {
                                    add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะใบงานไม่สามารถเป็น In Progress เนื่องจากผู้รับมอบหมายเป็นค่าว่าง");
                                    error_log = error_log + 1;
                                } else {
                                    casetaskGr.state = 2;
                                    casetaskGr.u_work_state = "เจ้าหน้าที่กำลังดำเนินการ";
                                }
							// เขียนเพิ่มจากเดิมไม่มี awaiting info
                            } else if(intt == "18" || intt == 18){
								casetaskGr.state = 18;
                                casetaskGr.u_work_state = "รอข้อมูลเพิ่มเติม";
							}

						}
						else if(sheet_state == "New" || sheet_state == "Open"){
							// value New:1, Open:10
							// if (sheet_result != "") {
							// 	error_log = error_log + 1;
							// 	add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด ผลการพิจารณาไม่สอดคล้องกับสถานะ");
							// }

							// var intt = get_select_state_active(sheet_state);
							// if (intt == "1" || intt == 1) {
                            //     casetaskGr.state = 1;
                            //     casetaskGr.u_work_state = "รอเจ้าหน้าที่เข้ามารับงาน";
                            // } else if(intt == "10" || intt == 10){
							// 	casetaskGr.state = 10;
                            //     casetaskGr.u_work_state = "รอเจ้าหน้าที่เข้ามารับงาน";
							// }

							add_log(source.sys_import_set, row_count, target_record, "Error", "สถานะใบงานไม่สามารถเป็น New หรือ Open ");
							error_log = error_log + 1;

						}
                        else if (sheet_state == "Resolved") {
                            // ผลการพิจารณาสามารเป็นเป็น  ["ได้ข้อสรุปกับลูกค้า","ไม่ได้ข้อสรุปกับลูกค้า"];
                            if (sheet_result == '') {
                                add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะใบงานไม่สามารถเป็น Resolved เนื่องจากผลการพิจารณาว่าง");
                                error_log = error_log + 1;
                            } else if ((sheet_product == '' || sheet_account_status == '' || sheet_contract_number == '' || sheet_total_debt == '' || sheet_principal == '') && (debt_project == "ทางด่วนแก้หนี้")) {
                                // ผลิตภัณฑ์, สถานะบัญชี, เลขที่สัญญา, ภาระหนี้รวม, เงินต้น
								add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะใบงานไม่สามารถเป็น Resolved เนื่องจากมี Mandatory Field ว่าง");
                                error_log = error_log + 1;
                            } else if (( sheet_product == '' || sheet_account_status == '' || sheet_contract_number == '' || sheet_principal == '') && (debt_project == "โครงการคุณสู้ เราช่วย")) {
                               // ผลิตภัณฑ์, สถานะบัญชี, เลขที่สัญญา, เงินต้น
								add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะใบงานไม่สามารถเป็น Resolved เนื่องจากมี Mandatory Field ว่าง");
                                error_log = error_log + 1;
                            } else if (check_select_state_active(sheet_state, sheet_result)) {
                                var resolution_code = get_active_result(sheet_result);
                                if (resolution_code == "Error") {
                                    add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะใบงาน");
                                    error_log = error_log + 1;
                                } else {
                                    casetaskGr.u_resolution_code = resolution_code;
                                    casetaskGr.state = 6;
                                    casetaskGr.u_work_state = "เจ้าหน้าที่ดำเนินการเสร็จสิ้น";
                                    casetaskGr.u_bulk_upload = true;
									casetaskGr.close_notes = sheet_additional_info;
                                }
                            } else {
                                    add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะใบงาน");
                                    error_log = error_log + 1;
                                }
                        }
                        else if (sheet_state == "Cancelled") {
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
                        }else if (sheet_state == "Closed") {
                            error_log = error_log + 1;
                            add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะใบงานไม่สามารถเป็น Closed ได้จากการ Duplicate ใบงาน");
                        }else{
                            error_log = error_log + 1;
                            add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด สถานะใบงานไม่ถูกต้อง");
                        }



						// แนวทางการช่วยเหลือลูกหนี้
						if (sheet_result == "ได้ข้อสรุปกับลูกค้า") {
							gs.info("duke แนวทางการช่วยเหลือลูกหนี้ insert");
							if (check_select_valid_activePipe2('u_glideline_debtor', "u_name=" + sheet_guidelines_debtors, project_debt_record)) {
								var glideline_debt = get_select_valid_active_check_fair('u_glideline_debtor', 'u_name', sheet_guidelines_debtors, project_debt_record);
								if (glideline_debt == "Error") {
									add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด แนวทางการช่วยเหลือลูกหนี้");
									error_log = error_log + 1;
								} else {

									casetaskGr.u_glideline_debt = glideline_debt;
								}
							} else {
								if (sheet_guidelines_debtors == "") {
									//add code
								} else {
									add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด แนวทางการช่วยเหลือลูกหนี้");
									error_log = error_log + 1;

								}
							}
						}


						// เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้ saveทับ
						if (sheet_result == "ไม่ได้ข้อสรุปกับลูกค้า") {
							if (check_select_valid_activePipe2('u_unable_help', "u_name=" + sheet_reason_not_help, project_debt_record)) {
								var reason_unable_help = get_select_valid_active_check_fair('u_unable_help', 'u_name', sheet_reason_not_help, project_debt_record);
								if (reason_unable_help == "Error") {
									add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้");
									error_log = error_log + 1;
								} else {
									casetaskGr.u_reason_unable_help = reason_unable_help;
								}

							} else {
								if (sheet_reason_not_help == "") {
									//add Code
								} else {
									add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้");
									error_log = error_log + 1;

								}
							}
						}
						// ภาระหนี้รวม (บาท)
						if (check_money_valid(sheet_total_debt)) {
							casetaskGr.u_debt_burden = set_money_object(sheet_total_debt);
						} else {
							if (sheet_total_debt == "") {
								//add code
							} else {
								add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด ภาระหนี้รวม (บาท)");
								error_log = error_log + 1;

							}
						}
						// เงินต้น (บาท)
						if (check_money_valid(sheet_principal)) {
							casetaskGr.u_principle = set_money_object(sheet_principal);
						} else {
							if (sheet_principal == "") {
								//add code
							} else {
								add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด เงินต้น (บาท)");
								error_log = error_log + 1;

							}
						}

						// ภาระหนี้ที่ตกลงชำระ (บาท)
						if (sheet_result == "ได้ข้อสรุปกับลูกค้า") {
							if (check_money_valid(sheet_debt_agreed)) {
								casetaskGr.u_debt_confirm = set_money_object(sheet_debt_agreed);
							} else {
								if (sheet_debt_agreed == "") {
									//add code
								} else {
									add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด ภาระหนี้ที่ตกลงชำระ (บาท)");
									error_log = error_log + 1;

								}
							}
						}

						// จำนวนงวดที่ชำระ (เดือน)
						if (sheet_result == "ได้ข้อสรุปกับลูกค้า") {
							try {
								casetaskGr.u_installment = parseInt(sheet_no_payment);
							} catch (e) {
								add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด ภาระหนี้ที่ตกลงชำระ (บาท)");
								error_log = error_log + 1;
							}
						}

						// ค่างวดต่อเดือน (บาท)
						if (sheet_result == "ได้ข้อสรุปกับลูกค้า") {
							if (check_money_valid(sheet_monthly)) {
								casetaskGr.u_amount_month = set_money_object(sheet_monthly);
							} else {
								if (sheet_monthly == "") {
									//add code
								} else {
									add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด ค่างวดต่อเดือน (บาท)");
									error_log = error_log + 1;

								}
							}
						}

						// รายงาน RDT
						if (check_date_valid(sheet_rtd_report)) {
							sheet_rtd_report = formatDate(sheet_rtd_report);
							if (check_date_valid(sheet_rtd_report)) {
								casetaskGr.u_rdt = set_date_object(sheet_rtd_report);
							} else {
								add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด รายงาน RDT");
								error_log = error_log + 1;
							}
						}

						// รายงาน DRD
						if (check_date_valid(sheet_drd_report)) {
							sheet_drd_report = formatDate(sheet_drd_report);
							if (check_date_valid(sheet_drd_report)) {
								casetaskGr.u_drd = set_date_object(sheet_drd_report);
							} else {
								add_log(source.sys_import_set, row_count, target_record, "Error", "ผิดพลาด วันที่ทำสัญญา");
								error_log = error_log + 1;
							}
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

							// gs.info('rowCount' + rowCount);

							var type = 'casetask';
							var recordNumber = new global.MF_ChildRecordNumberPadding().padNumber(rowCountNumber, parentNumber, type);
							// source.u_หมายเลข_case_task = recordNumber;
							gs.info('recordNumber' + recordNumber);
							// source.update();

							// countNumber++;
							casetaskGr.number = recordNumber;
							casetaskGr.insert();

							// test(targetGR, recordNumber);
							var case_number = recordNumber;
							add_log(source.sys_import_set, row_count, case_number, "Insert","เพิ่มข้อมูลสำเร็จ");
						} else {
							error_log_all = 1;
							add_log(source.sys_import_set, row_count, target_record, "Error", "เพิ่มข้อมูลไม่สำเร็จ มีข้อผิดพลาด " + String(error_log) + " ข้อ");

						}

					} else {
						//ไม่เจอ Case
						add_log(source.sys_import_set, row_count, target_record, "Error", "ค้นหา Case ID ไม่เจอ! " + sheet_case_relate);
						error_log = error_log + 1;
						error_log_all = 1;

					}
				} else {
					add_log(source.sys_import_set, row_count, target_record, "Error", "ผิด Format การอัพเดตข้อมูล ");
					error_log = error_log + 1;
					error_log_all = 1;


				}
			}
		}
		gs.info("[Excel Bulk Upload] Start Countung");
		sleep(5000);

		var log_total = 0;
		var log_error_total = 0;
		var selectGrRecord = new GlideRecord("x_baot_debt_sett_0_import_excel_logs");
		selectGrRecord.addQuery("importsetref", source.sys_import_set);
		selectGrRecord.query();
		if (selectGrRecord.next()) {
			log_total = selectGrRecord.getRowCount();
		}
		gs.info("[Excel Bulk Upload]source.sys_import_set : "+source.sys_import_set);
		var erorrGrRecord = new GlideRecord("x_baot_debt_sett_0_import_excel_logs");
		erorrGrRecord.addQuery("importsetref", source.sys_import_set);
		erorrGrRecord.addQuery("state", 'ERROR');
		erorrGrRecord.query();
		if (erorrGrRecord.next()) {
			log_error_total = erorrGrRecord.getRowCount();
		}


		// Update Import Set Ref in Import Attachment Table


		//importExcelAttachmentSysId มืออยู่แล้ว



		// gs.info("[Excel Bulk Upload]source.sys_import: "+source.sys_import_set.data_source.sys_id);
		// var gr = new GlideRecord('sys_attachment');
		// gr.addQuery('table_name', 'sys_data_source');
		// gr.addQuery('table_sys_id', source.sys_import_set.data_source.sys_id);
		// gr.orderByDesc('sys_created_on');
		// gr.query();
		// if (gr.next()) {
			// Update Import Set in Table Import Attachment
			// var sysid_table_importattachment = gr.getValue("file_name").split("-")[0]; // ไม่ต้องใช้เพราะว่าประกาศไว้แล้วข้างบนคือ importExcelAttachmentSysId


			var row_table_importattachment = new GlideRecord('x_baot_debt_sett_0_import_excel_attachments');
			row_table_importattachment.addQuery("sys_id", importExcelAttachmentSysId);
			row_table_importattachment.query();
			if (row_table_importattachment.next()) {
				// find Import Set

				// var row_importset = new GlideRecord('sys_import_set');
				// row_importset.addQuery("sys_id", import_set.sys_id);
				// row_importset.query();
				// if (row_importset.next()) {

					// gs.info("[Excel Bulk Upload] Import Set add(sys_import_set): " + import_set.sys_id + " | " + row_importset.sys_id+" | "+ source.sys_import_set);
					// gs.info("[Excel Bulk Upload] importsetLink "+row_importset.sys_id);

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

				// } else {
				// 	gs.info("[Excel Bulk Upload] Import Set not found for sys_id(sys_import_set): " + import_set.sys_id);
				// 	row_table_importattachment.setValue("state", "5");
				// }

			} else { 
				// ตัวนี้ต้องมี
				gs.info("[Excel Bulk Upload] Record row_table_importattachment not found!");
			}
		// } //ของ sys attachment



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


	const regex = /^(\d{2})-(\d{2})-(\d{4}) (\d{2}):(\d{2})$/;
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

	const day = parseInt(match2[1], 10);
	const month = parseInt(match2[2], 10);
	const year = parseInt(match2[3], 10);
	const hour = parseInt(match2[4], 10);
	const minute = parseInt(match2[5], 10);


	if (day < 1 || day > 31 || month < 1 || month > 12 || hour < 0 || hour > 23 || minute < 0 || minute > 59) {
		return false;
	}

	const daysInMonth = [31, (year % 4 == 0 && year % 100 !== 0) || (year % 400 == 0) ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
	if (day > daysInMonth[month - 1]) {
		return false;
	}

	return true;
}

function set_date_object(date_string_input) {
	const regex = /^(\d{2})-(\d{2})-(\d{4})$/;
	const match = date_string_input.match(regex);
	const day = parseInt(match[1], 10);
	const month = parseInt(match[2], 10);
	const year = parseInt(match[3], 10);

	return year + "-" + month + "-" + day;
}

function check_date_valid(date_input) {
	const regex = /^(\d{2})-(\d{2})-(\d{4})$/;
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

function check_idcard_valid(id_number) {
	const regex = /^[0-9]{13}$/;
	return regex.test(id_number);
}

function set_money_object(money_input) {
	var x = money_input.replace("฿", "");
	var y = x.replace(",", "");
	return "THB;" + y;
}

function check_money_valid(amount) {
	const regex = /^(฿)?(\d{1,3}(,\d{3})*|\d+)(\.\d{1,2})?$/;
	return regex.test(amount);
}

function check_email_valid(email) {
	const regex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
	return regex.test(email);
}

function check_phone_valid(phoneNumber) {
	gs.info('phoneNumber' + phoneNumber);
	const regex = /^66\d{9}$/;
	var text = String(phoneNumber).replace(" ", "");
	gs.info('text' + text);
	return regex.test(text);

}

function check_number_valid(inputNumber) {
	const regex = /^\d+$/;
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
	gs.info("[Excel Bulk Upload IsNotMember]user: "+ userSysId+" group: "+groupSysId);
	return memberOfGroup;
}

function sleep(ms) {
	var unixtime_ms = new Date().getTime();
	while (new Date().getTime() < unixtime_ms + ms) { }

}


function test(targetGR, sheet_case_task_id) {
	//หมายเลข Case 1
	var sheet_case_id = targetGR.u_หมายเลข_case;
	// หมายเลข Case Task 2
	// var sheet_case_task_id = targetGR.u_หมายเลข_case_task;
	// case ที่เกี่ยวข้อง 3
	var sheet_case_relate = targetGR.u_case_ท__เก__ยวข_อง;
	//case detail 4
	var sheet_case_detail = targetGR.u_หมายเลข_detail;
	// เรื่อง 5
	var sheet_title = targetGR.u_เร__อง;
	// สถานะการดำเนินงาน 6
	var sheet_status = targetGR.u_สถานะการดำเน_นงาน;
	// SLA Due date 7
	var sheet_sla_duedate = targetGR.u_sla_due_date;
	// กลุ่มที่รับผิดชอบ 8
	var sheet_responsible_group = targetGR.u_กล__มท__ร_บผ_ดชอบ;
	// Due date 9
	var sheet_duedate = targetGR.u_due_date;
	// ผู้รับมอบหมาย 10
	var sheet_assignee = targetGR.u_ผ__ร_บมอบหมาย;
	// ขอรับบริการในนาม 11
	var sheet_request_service_on_behalf_of = targetGR.u_ขอร_บบร_การในนาม;
	// ผู้ติดต่อ 12
	var sheet_contact = targetGR.u_ผ__ต_ดต_อ;
	// ผู้ขอรับบริการ 13
	var sheet_service_requester = targetGR.u_ผ__ขอร_บบร_การ;
	// ประเภทเลขที่ยืนยันตัวตน 14
	var sheet_type_identification_number = targetGR.u_ประเภทเลขท__ย_นย_นต_วตน;
	// ประเภทเลขที่ยืนยันตัวตน (นิติบุคคล) 15
	var sheet_type_identification_number_juristic = targetGR.u_ประเภทเลขท___ตน__น_ต_บ_คคล_;
	// เลขที่ยืนยันตัวตน 16
	var sheet_identification_number = targetGR.u_เลขท__ย_นย_นต_วตน;
	// เลขที่ยืนยันตัวตน (นิติบุคคล) 17
	var sheet_identification_number_juristic = targetGR.u_เลขท__ย_นย_น_ตน__น_ต_บ_คคล_;
	// หมายเลขโทรศัพท์ 18
	var sheet_phone = targetGR.u_หมายเลขโทรศ_พท_;
	// อีเมล 19
	var sheet_email = targetGR.u_อ_เมล;
	// หมายเลขโทรศัพท์ (สำรอง) 20
	var sheet_phone_backup = targetGR.u_หมายเลขโทรศ_พท___สำรอง_;
	// อีเมล (สำรอง) 21
	var sheet_email_backup = targetGR.u_อ_เมล__สำรอง_;
	// จังหวัด (ที่อยู่ลูกหนี้) 22
	var sheet_province = targetGR.u_จ_งหว_ด__ท__อย__ล_กหน___;
	// ผู้ให้บริการ 23
	var sheet_provider = targetGR.u_ผ__ให_บร_การ;
	// ผลิตภัณฑ์ 24
	var sheet_product = targetGR.u_ผล_ตภ_ณฑ_;
	if (sheet_product == '' || sheet_product == null) {
		//    gs.info('sheet_guidelines_debtorssssssssssssllllllllllds');

	} else {
		sheet_product = targetGR.u_ผล_ตภ_ณฑ_.split("|")[0].trim();
	}
	// สถานะบัญชี 25
	var sheet_account_status = targetGR.u_สถานะบ_ญช_;
	if (sheet_account_status == '' || sheet_account_status == null) {
		//    gs.info('sheet_guidelines_debtorssssssssssssllllllllllds');
	} else {
		sheet_account_status = targetGR.u_สถานะบ_ญช_.split("|")[0].trim();
	}
	//โครงการแก้หนี้ 26
	var sheet_project_debt = targetGR.u_โครงการแก_หน__;
	// เลขที่ทะเบียนรถ 27
	var sheet_vehicle_number = targetGR.u_เลขท__ทะเบ_ยนรถ;
	// จังหวัดที่จดทะเบียนรถ 28
	var sheet_vehicle_number_province = targetGR.u_จ_งหว_ดท__จดทะเบ_ยนรถ;
	// แนวทางที่ต้องการเสนอให้ผู้บริการทางการเงินพิจารณา 29
	var sheet_guidelines_provider = targetGR.u_แนวทางท__ต_อ_การเง_นพ_จารณา;
	// เหตุผลหรือคำอธิบายประกอบคำขอ 30
	var sheet_reason = targetGR.u_เหต_ผลหร_อคำอธ_บายประกอบคำขอ;
	// // Case ที่เกี่ยวข้อง 31
	var sheet_case_relate_2 = targetGR.u_case_ท__เก__ยวข_อง;
	// Ref Case Task 32
	var sheet_case_ref = targetGR.u_ref_case_task;
	// เลขที่บัตร/ เลขที่สัญญา 33
	var sheet_contract_number = targetGR.u_เลขท__บ_ตร__เลขท__ส_ญญา;
	// วันที่เริ่มติดต่อลูกหนี้ 34
	var sheet_date_of_start = targetGR.u_ว_นท__เร__มต_ดต_อล_กหน__;
	// ผลการพิจารณา 35
	var sheet_result = targetGR.u_ผลการพ_จารณา;
	// แนวทางการช่วยเหลือลูกหนี้ 36
	var sheet_guidelines_debtors = targetGR.u_แนวทางการช_วยเหล_อล_กหน__;
	if (sheet_guidelines_debtors == '' || sheet_guidelines_debtors == null) {
		//gs.info('sheet_guidelines_debtorssssssssssssllllllllllds');
	} else {
		//  gs.info('sheet_guidelines_debtorssssssssseeeeeeeeeeeeeeeedlds');
		sheet_guidelines_debtors = targetGR.u_แนวทางการช_วยเหล_อล_กหน__.split("|")[0].trim();
	}
	// เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้ 37
	var sheet_reason_not_help = targetGR.u_เหต_ผลท__ไม__หล_อล_กหน__ได_;
	if (sheet_reason_not_help == '' || sheet_reason_not_help == null) {
		//    gs.info('sheet_guidelines_debtorssssssssssssllllllllllds');
	} else {
		sheet_reason_not_help = targetGR.u_เหต_ผลท__ไม__หล_อล_กหน__ได_.split("|")[0].trim();
	}
	// ภาระหนี้รวม (บาท) 38
	var sheet_total_debt = targetGR.u_ภาระหน__รวม__บาท_;
	// เงินต้น (บาท) 39
	var sheet_principal = targetGR.u_เง_นต_น__บาท_;
	// ภาระหนี้ที่ตกลงชำระ (บาท) 40
	var sheet_debt_agreed = targetGR.u_ภาระหน__ท__ตกลงชำระ__บาท_;
	// จำนวนงวดที่ชำระ (เดือน) 41
	var sheet_no_payment = targetGR.u_จำนวนงวดท__ชำระ__เด_อน_;
	// ค่างวดต่อเดือน (บาท) 42
	var sheet_monthly = targetGR.u_ค_างวดต_อเด_อน__บาท_;
	// รายงาน RDT 43
	var sheet_rtd_report = targetGR.u_รายงาน_rdt;
	sheet_rtd_report = formatDate(sheet_rtd_report);
	// รายงาน DRD 44
	var sheet_drd_report = targetGR.u_รายงาน_drd ? targetGR.u_รายงาน_drd : targetGR.u_ว_นท__ทำส_ญญา;
	sheet_drd_report = formatDate(sheet_drd_report);
	// ข้อมูลที่ต้องการแจ้ง ธปท. เพิ่มเติม (ถ้ามี) 45
	var sheet_additional_info = targetGR.u_ข_อม_ลท__ต_อ__มเต_ม__ถ_าม__;
	// สถานะใบงาน 46
	var sheet_state = targetGR.u_สถานะใบงาน;
	// ประเภทการแก้ไข 47
	var sheet_cancel_state = targetGR.u_ประเภทการแก_ไข;
	// close note 48
	var sheet_close = targetGR.u_close_note;


	var casetaskGr = new GlideRecord('x_baot_debt_sett_0_ect_debt_settlement');
	casetaskGr.addQuery('number', sheet_case_task_id);
	casetaskGr.query();
	if (casetaskGr.next()) {
		// มีทั้งหมด 48 ฟิลด์

		// หมายเลข Case 1
		if (casetaskGr.u_display_number.getDisplayValue() == sheet_case_id) {
			gs.info('[Excel Bulk Upload Test Pass]หมายเลข Case is equal');
		} else {
			gs.info('[Excel Bulk Upload Test Fail]หมายเลข Case is not equal| u_display_number == sheet_case_id : ' + casetaskGr.u_display_number.getDisplayValue() + ' == ' + sheet_case_id);
		}
		// หมายเลข Case Task 2
		if (casetaskGr.number.getDisplayValue() == sheet_case_task_id) {
			gs.info('[Excel Bulk Upload Test Pass]หมายเลข Case Task is equal');
		} else {
			gs.info('[Excel Bulk Upload Test Fail]หมายเลข Case Task is not equal| number == sheet_case_task_id : ' + casetaskGr.number.getDisplayValue() + ' == ' + sheet_case_task_id);
		}
		// case ที่เกี่ยวข้อง 3
		if (casetaskGr.u_major_case.getDisplayValue() == sheet_case_relate) {
			gs.info('[Excel Bulk Upload Test Pass]case ที่เกี่ยวข้อง is equal');
		} else {
			gs.info('[Excel Bulk Upload Test Fail]case ที่เกี่ยวข้อง is not equal| u_major_case == sheet_case_relate : ' + casetaskGr.u_major_case.getDisplayValue() + ' == ' + sheet_case_relate);
		}
		//case detail 4
		// not use
		// เรื่อง 5
		if (casetaskGr.short_description == sheet_title) {
			gs.info('[Excel Bulk Upload Test Pass]เรื่อง is equal');
		} else {
			gs.info('[Excel Bulk Upload Test Fail]เรื่อง is not equal| short_description == sheet_title : ' + casetaskGr.short_description + ' == ' + sheet_title);
		}
		// สถานะการดำเนินงาน 6
		// ไม่ถูกเรียกใช้ u_work_state = 
		// SLA Due date 7
		if (casetaskGr.u_sla_due_date.getDisplayValue() == sheet_sla_duedate) {
			gs.info('[Excel Bulk Upload Test Pass]SLA Due date is equal');
		} else {
			gs.info('[Excel Bulk Upload Test Fail]SLA Due date is not equal| u_sla_due_date == sheet_sla_duedate : ' + casetaskGr.u_sla_due_date.getDisplayValue() + ' == ' + sheet_sla_duedate);
		}
		// กลุ่มที่รับผิดชอบ 8
		if (casetaskGr.assignment_group.getDisplayValue() == sheet_responsible_group) {
			gs.info('[Excel Bulk Upload Test Pass]กลุ่มที่รับผิดชอบ is equal');
		} else {
			gs.info('[Excel Bulk Upload Test Fail]กลุ่มที่รับผิดชอบ is not equal| assignment_group == sheet_responsible_group : ' + casetaskGr.assignment_group.getDisplayValue() + ' == ' + sheet_responsible_group);
		}
		// Due date 9
		if (casetaskGr.due_date.getDisplayValue() == sheet_duedate) {
			gs.info('[Excel Bulk Upload Test Pass]Due date is equal');
		} else {
			gs.info('[Excel Bulk Upload Test Fail]Due date is not equal| due_date == sheet_duedate : ' + casetaskGr.due_date.getDisplayValue() + ' == ' + sheet_duedate);
		}
		// ผู้รับมอบหมาย 10
		if (casetaskGr.assigned_to.getDisplayValue() == sheet_assignee) {
			gs.info('[Excel Bulk Upload Test Pass]ผู้รับมอบหมาย is equal');
		} else {
			gs.info('[Excel Bulk Upload Test Fail]ผู้รับมอบหมาย is not equal| assigned_to == sheet_assignee : ' + casetaskGr.assigned_to.getDisplayValue() + ' == ' + sheet_assignee);
		}
		// ขอรับบริการในนาม 11
		//ไม่ถูกใช้งาน
		// ผู้ติดต่อ 12
		//ไม่ถูกใช้งาน 
		// ผู้ขอรับบริการ 13
		if (casetaskGr.u_consumer.getDisplayValue() == sheet_service_requester) {
			gs.info('[Excel Bulk Upload Test Pass]ผู้ขอรับบริการ is equal');
		} else {
			gs.info('[Excel Bulk Upload Test] Failผู้ขอรับบริการ is not equal| u_consumer == sheet_service_requester : ' + casetaskGr.u_consumer.getDisplayValue() + ' == ' + sheet_service_requester);
		}
		// ประเภทเลขที่ยืนยันตัวตน 14
		//ไม่ถูกใช้งาน
		// ประเภทเลขที่ยืนยันตัวตน (นิติบุคคล) 15
		//ไม่ถูกใช้งาน
		// เลขที่ยืนยันตัวตน 16
		//ไม่ถูกใช้งาน
		// เลขที่ยืนยันตัวตน (นิติบุคคล) 17
		// หมายเลขโทรศัพท์ 18
		// ไม่ถูกใช้งาน
		// อีเมล 19
		// ไม่ถูกใช้งาน
		// หมายเลขโทรศัพท์ (สำรอง) 20
		if (casetaskGr.u_secondary_phone.getDisplayValue() == sheet_phone_backup) {
			gs.info('[Excel Bulk Upload Test Pass]หมายเลขโทรศัพท์ (สำรอง) is equal');
		} else {
			gs.info('[Excel Bulk Upload Test Fail]หมายเลขโทรศัพท์ (สำรอง) is not equal| u_secondary_phone == sheet_phone_backup : ' + casetaskGr.u_secondary_phone.getDisplayValue() + ' == ' + sheet_phone_backup);
		}
		// อีเมล (สำรอง) 21
		if (casetaskGr.u_secondary_email.getDisplayValue() == sheet_email_backup) {
			gs.info('[Excel Bulk Upload Test Pass]อีเมล (สำรอง) is equal');
		} else {
			gs.info('[Excel Bulk Upload Test Fail]อีเมล (สำรอง) is not equal| u_secondary_email == sheet_email_backup : ' + casetaskGr.u_secondary_email.getDisplayValue() + ' == ' + sheet_email_backup);
		}
		// จังหวัด (ที่อยู่ลูกหนี้) 22
		//ไม่ถูกใช้งาน
		// ผู้ให้บริการ 23
		if (casetaskGr.u_provider.getDisplayValue() == sheet_provider) {
			gs.info('[Excel Bulk Upload Test Pass]ผู้ให้บริการ is equal');
		} else {
			gs.info('[Excel Bulk Upload Test Fail]ผู้ให้บริการ is not equal| u_provider == sheet_provider : ' + casetaskGr.u_provider.getDisplayValue() + ' == ' + sheet_provider);
		}
		// ผลิตภัณฑ์ 24
		if (casetaskGr.u_product.getDisplayValue() == sheet_product) {
			gs.info('[Excel Bulk Upload Test Pass]ผลิตภัณฑ์ is equal');
		} else {
			gs.info('[Excel Bulk Upload Test Fail]ผลิตภัณฑ์ is not equal| u_product == sheet_product : ' + casetaskGr.u_product.getDisplayValue() + ' == ' + sheet_product);
		}
		// สถานะบัญชี 25
		if (casetaskGr.u_state_debt.getDisplayValue() == sheet_account_status) {
			gs.info('[Excel Bulk Upload Test Pass]สถานะบัญชี is equal');
		} else {
			gs.info('[Excel Bulk Upload Test Fail]สถานะบัญชี is not equal| u_state_debt == sheet_account_status : ' + casetaskGr.u_state_debt.getDisplayValue() + ' == ' + sheet_account_status);
		}
		//โครงการแก้หนี้ 26
		if (casetaskGr.u_debt_project.getDisplayValue() == sheet_project_debt) {
			gs.info('[Excel Bulk Upload Test Pass]โครงการแก้หนี้ is equal');
		} else {
			gs.info('[Excel Bulk Upload Test Fail]โครงการแก้หนี้ is not equal| u_debt_project == sheet_project_debt : ' + casetaskGr.u_debt_project.getDisplayValue() + ' == ' + sheet_project_debt);
		}
		// เลขที่ทะเบียนรถ 27
		if (casetaskGr.u_number_car == sheet_vehicle_number) {
			gs.info('[Excel Bulk Upload Test Pass]เลขที่ทะเบียนรถ is equal');
		} else {
			gs.info('[Excel Bulk Upload Test Fail]เลขที่ทะเบียนรถ is not equal| u_number_car == sheet_vehicle_number : ' + casetaskGr.u_number_car + ' == ' + sheet_vehicle_number);
		}
		// จังหวัดที่จดทะเบียนรถ 28
		if (casetaskGr.u_province_car.getDisplayValue() == sheet_vehicle_number_province) {
			gs.info('[Excel Bulk Upload Test Pass]จังหวัดที่จดทะเบียนรถ is equal');
		} else {
			gs.info('[Excel Bulk Upload Test Fail]จังหวัดที่จดทะเบียนรถ is not equal| u_province_car == sheet_vehicle_number_province : ' + casetaskGr.u_province_car.getDisplayValue() + ' == ' + sheet_vehicle_number_province);
		}
		// แนวทางที่ต้องการเสนอให้ผู้บริการทางการเงินพิจารณา 29
		if (casetaskGr.u_offer_provider.getDisplayValue() == sheet_guidelines_provider) {
			gs.info('[Excel Bulk Upload Test Pass]แนวทางที่ต้องการเสนอให้ผู้บริการทางการเงินพิจารณา is equal');
		} else {
			gs.info('[Excel Bulk Upload Test Fail]แนวทางที่ต้องการเสนอให้ผู้บริการทางการเงินพิจารณา is not equal| u_offer_provider == sheet_guidelines_provider : ' + casetaskGr.u_offer_provider.getDisplayValue() + ' == ' + sheet_guidelines_provider);
		}
		// เหตุผลหรือคำอธิบายประกอบคำขอ 30
		if (casetaskGr.u_provider_verified == sheet_reason) {
			gs.info('[Excel Bulk Upload Test Pass]เหตุผลหรือคำอธิบายประกอบคำขอ is equal');
		} else {
			gs.info('[Excel Bulk Upload Test Fail]เหตุผลหรือคำอธิบายประกอบคำขอ is not equal| u_provider_verified == sheet_reason : ' + casetaskGr.u_provider_verified + ' == ' + sheet_reason);
		}
		// Case ที่เกี่ยวข้อง 31
		if (casetaskGr.u_major_case.getDisplayValue() == sheet_case_relate_2) {
			gs.info('[Excel Bulk Upload Test Pass]Case ที่เกี่ยวข้อง is equal');
		} else {
			gs.info('[Excel Bulk Upload Test Fail]Case ที่เกี่ยวข้อง is not equal| u_major_case == sheet_case_relate_2 : ' + casetaskGr.u_major_case.getDisplayValue() + ' == ' + sheet_case_relate_2);
		}
		// Ref Case Task 32
		//อันนี้ไม่ได้เอาไป update
		// เลขที่บัตร/ เลขที่สัญญา 33
		if (casetaskGr.u_number_contact == sheet_contract_number) {
			gs.info('[Excel Bulk Upload Test Pass]เลขที่บัตร/ เลขที่สัญญา is equal');
		} else {
			gs.info('[Excel Bulk Upload Test Fail]เลขที่บัตร/ เลขที่สัญญา is not equal| u_number_contact == sheet_contract_number : ' + casetaskGr.u_number_contact + ' == ' + sheet_contract_number);
		}
		// วันที่เริ่มติดต่อลูกหนี้ 34
		if (casetaskGr.u_contact_debt.getDisplayValue() == sheet_date_of_start) {
			gs.info('[Excel Bulk Upload Test Pass]วันที่เริ่มติดต่อลูกหนี้ is equal');
		} else {
			gs.info('[Excel Bulk Upload Test Fail]วันที่เริ่มติดต่อลูกหนี้ is not equal| u_contact_debt == sheet_date_of_start : ' + casetaskGr.u_contact_debt.getDisplayValue() + ' == ' + sheet_date_of_start);
		}
		// ผลการพิจารณา 35
		if (casetaskGr.u_resolution_code.getDisplayValue() == sheet_result) {
			gs.info('[Excel Bulk Upload Test Pass]ผลการพิจารณา is equal');
		} else {
			gs.info('[Excel Bulk Upload Test Fail]ผลการพิจารณา is not equal| u_resolution_code == sheet_result : ' + casetaskGr.u_resolution_code.getDisplayValue() + ' == ' + sheet_result);
		}
		// แนวทางการช่วยเหลือลูกหนี้ 36
		if (casetaskGr.u_glideline_debt.getDisplayValue() == sheet_guidelines_debtors) {
			gs.info('[Excel Bulk Upload Test Pass]แนวทางการช่วยเหลือลูกหนี้ is equal');
		} else {
			gs.info('[Excel Bulk Upload Test Fail]แนวทางการช่วยเหลือลูกหนี้ is not equal| u_glideline_debt == sheet_guidelines_debtors : ' + casetaskGr.u_glideline_debt.getDisplayValue() + ' == ' + sheet_guidelines_debtors);
		}
		// เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้ 37
		if (casetaskGr.u_reason_unable_help.getDisplayValue() == sheet_reason_not_help) {
			gs.info('[Excel Bulk Upload Test Pass]เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้ is equal');
		} else {
			gs.info('[Excel Bulk Upload Test Fail]เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้ is not equal| u_reason_unable_help == sheet_reason_not_help : ' + casetaskGr.u_reason_unable_help.getDisplayValue() + ' == ' + sheet_reason_not_help);
		}
		// ภาระหนี้รวม (บาท) 38
		if (casetaskGr.u_debt_burden.getDisplayValue() == sheet_total_debt) {
			gs.info('[Excel Bulk Upload Test Pass]ภาระหนี้รวม (บาท) is equal');
		} else {
			gs.info('[Excel Bulk Upload Test Fail]ภาระหนี้รวม (บาท) is not equal| u_debt_burden == sheet_total_debt : ' + casetaskGr.u_debt_burden.getDisplayValue() + ' == ' + sheet_total_debt);
		}
		// เงินต้น (บาท) 39
		if (casetaskGr.u_principle.getDisplayValue() == sheet_principal) {
			gs.info('[Excel Bulk Upload Test Pass]เงินต้น (บาท) is equal');
		} else {
			gs.info('[Excel Bulk Upload Test Fail]เงินต้น (บาท) is not equal| u_principle == sheet_principal : ' + casetaskGr.u_principle.getDisplayValue() + ' == ' + sheet_principal);
		}
		// ภาระหนี้ที่ตกลงชำระ (บาท) 40
		if (casetaskGr.u_debt_confirm.getDisplayValue() == sheet_debt_agreed) {
			gs.info('[Excel Bulk Upload Test Pass]ภาระหนี้ที่ตกลงชำระ (บาท) is equal');
		} else {
			gs.info('[Excel Bulk Upload Test Fail]ภาระหนี้ที่ตกลงชำระ (บาท) is not equal| u_debt_confirm == sheet_debt_agreed : ' + casetaskGr.u_debt_confirm.getDisplayValue() + ' == ' + sheet_debt_agreed);
		}
		// จำนวนงวดที่ชำระ (เดือน) 41
		if (casetaskGr.u_installment == sheet_no_payment) {
			gs.info('[Excel Bulk Upload Test Pass]จำนวนงวดที่ชำระ (เดือน) is equal');
		} else {
			gs.info('[Excel Bulk Upload Test Fail]จำนวนงวดที่ชำระ (เดือน) is not equal| u_installment == sheet_no_payment : ' + casetaskGr.u_installment + ' == ' + sheet_no_payment);
		}
		// ค่างวดต่อเดือน (บาท) 42
		if (casetaskGr.u_amount_month.getDisplayValue() == sheet_monthly) {
			gs.info('[Excel Bulk Upload Test Pass]ค่างวดต่อเดือน (บาท) is equal');
		} else {
			gs.info('[Excel Bulk Upload Test Fail]ค่างวดต่อเดือน (บาท) is not equal| u_amount_month == sheet_monthly : ' + casetaskGr.u_amount_month.getDisplayValue() + ' == ' + sheet_monthly);
		}
		// รายงาน RDT 43
		if (casetaskGr.u_rdt.getDisplayValue() == sheet_rtd_report) {
			gs.info('[Excel Bulk Upload Test Pass]รายงาน RDT is equal');
		} else {
			gs.info('[Excel Bulk Upload Test Fail]รายงาน RDT is not equal| u_rdt == sheet_rtd_report : ' + casetaskGr.u_rdt.getDisplayValue() + ' == ' + sheet_rtd_report);
		}
		// รายงาน DRD 44
		if (casetaskGr.u_drd.getDisplayValue() == sheet_drd_report) {
			gs.info('[Excel Bulk Upload Test Pass]รายงาน DRD is equal');
		} else {
			gs.info('[Excel Bulk Upload Test Fail]รายงาน DRD is not equal| u_drd == sheet_drd_report : ' + casetaskGr.u_drd.getDisplayValue() + ' == ' + sheet_drd_report);
		}
		// ข้อมูลที่ต้องการแจ้ง ธปท. เพิ่มเติม (ถ้ามี) 45
		if (casetaskGr.u_other_informantion == sheet_additional_info) {
			gs.info('[Excel Bulk Upload Test Pass]ข้อมูลที่ต้องการแจ้ง ธปท. เพิ่มเติม (ถ้ามี) is equal');
		} else {
			gs.info('[Excel Bulk Upload Test Fail]ข้อมูลที่ต้องการแจ้ง ธปท. เพิ่มเติม (ถ้ามี) is not equal| u_close_notes == sheet_additional_info : ' + casetaskGr.u_other_informantion + ' == ' + sheet_additional_info);
		}
		// สถานะใบงาน 46
		if (casetaskGr.state.getDisplayValue() == sheet_state) {
			gs.info('ส[Excel Bulk Upload Test Pass]ถานะใบงาน is equal');
		} else {
			gs.info('[Excel Bulk Upload Test Fail]สถานะใบงาน is not equal| state == sheet_state : ' + casetaskGr.state.getDisplayValue() + ' == ' + sheet_state);
		}
		// ประเภทการแก้ไข 47
		//ไม่ได้ใช้งาน
		// close note 48
		if (casetaskGr.close_notes == sheet_close) {
			gs.info('[Excel Bulk Upload Test Pass]close note is equal');
		} else {
			gs.info('[Excel Bulk Upload Test Fail]close note is not equal| close_notes == sheet_close : ' + casetaskGr.close_notes + ' == ' + sheet_close);
		}

		//จำนวนฟิลด์ที่ถูกใช้งานได้ 35 ฟิลด์

	}



}