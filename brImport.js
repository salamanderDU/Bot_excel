(function executeRule(current, previous /*null when async*/) {
	var attachmentApi = new GlideSysAttachment();

	var attachmentGR = new GlideRecord('sys_attachment');
	attachmentGR.addQuery('table_sys_id', current.sys_id);
	attachmentGR.query();
	if (attachmentGR.next()) {
		if (attachmentGR.getRowCount() != 1) {
			current.state = "5";
			gs.addInfoMessage("ไม่สามารถอัปโหลดไฟล์ excel มากกว่า 1 ไฟล์ได้");
			current.u_file_error = "ไม่สามารถอัปโหลดไฟล์ excel มากกว่า 1 ไฟล์ได้";

		} else {

			//เข้าหา header ในไฟล์ excel
			var sysId = attachmentGR.sys_id;
			var column_type = checkColumn(current, sysId); // this function insert log error on filed.

			var originalFilename = attachmentGR.file_name.toString();

			if ((!originalFilename.startsWith(current.sys_id) && column_type)) {
				var newFilename = current.sys_id + '-' + originalFilename;
				var contentStream = attachmentApi.getContentStream(attachmentGR.getUniqueValue());
				attachmentApi.writeContentStream(current, newFilename, attachmentGR.content_type, contentStream);
				attachmentApi.deleteAttachment(attachmentGR.getUniqueValue());

				current.filename = originalFilename;
				var userSysId = gs.getUserID();
				current.upload_by = userSysId;

				var newNameFile;
				var attachmentGR1 = new GlideRecord('sys_attachment');
				attachmentGR1.addQuery('table_sys_id', current.sys_id);
				attachmentGR1.query();
				if (attachmentGR1.next()) {
					newNameFile = attachmentGR1.file_name.toString();
				}

				if (checkExcelFile(newNameFile)) {
					selectState(current, column_type);
				}
				else {
					current.state = "5";
					gs.addInfoMessage("ชื่อไฟล์ excel ผิดพลาด!");
					current.u_file_error = "ชื่อไฟล์ excel ผิดพลาด!";
				}

			}
		}
	} else {

		gs.addInfoMessage("ไม่พบ Attachment ");
		current.state = "5";
		current.u_file_error = "ไม่พบ Attachment ";
	}

})(current, previous);

function selectState(current, column_type) {
	if (column_type == 1) {
		current.state = 6;
	} else if (column_type == 2) {
		current.state = 7;
	}
}

function checkColumn(current, sysId) {
	var reqCol_old = [
		"หมายเลข Case Task",
		"เรื่อง",
		"สถานะการดำเนินงาน",
		"SLA Due date",
		"กลุ่มที่รับผิดชอบ",
		"Due date",
		"ผู้รับมอบหมาย",
		"หมายเลข Case",
		"ขอรับบริการในนาม",
		"ผู้ติดต่อ",
		"ผู้ขอรับบริการ",
		"ประเภทเลขที่ยืนยันตัวตน",
		"ประเภทเลขที่ยืนยันตัวตน (นิติบุคคล)",
		"เลขที่ยืนยันตัวตน",
		"เลขที่ยืนยันตัวตน (นิติบุคคล)",
		"หมายเลขโทรศัพท์",
		"อีเมล",
		"หมายเลขโทรศัพท์ (สำรอง)",
		"อีเมล (สำรอง)",
		"จังหวัด (ที่อยู่ลูกหนี้)",
		"ชื่อนิติบุคคล",
		"ผู้ให้บริการ",
		"ผลิตภัณฑ์",
		"สถานะบัญชี",
		"โครงการแก้หนี้",
		"เลขที่ทะเบียนรถ",
		"จังหวัดที่จดทะเบียนรถ",
		"แนวทางที่ต้องการเสนอให้ผู้บริการทางการเงินพิจารณา",
		"เหตุผลหรือคำอธิบายประกอบคำขอ",
		"Case ที่เกี่ยวข้อง",
		"Ref Case Task",
		"เลขที่บัตร/ เลขที่สัญญา",
		"วันที่เริ่มติดต่อลูกหนี้",
		"ผลการพิจารณา",
		"แนวทางการช่วยเหลือลูกหนี้",
		"เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้",
		"ภาระหนี้รวม (บาท)",
		"เงินต้น (บาท)",
		"ภาระหนี้ที่ตกลงชำระ (บาท)",
		"จำนวนงวดที่ชำระ (เดือน)",
		"ค่างวดต่อเดือน (บาท)",
		"รายงาน RDT",
		"วันที่ทำสัญญา",
		"ข้อมูลที่ต้องการแจ้ง ธปท. เพิ่มเติม (ถ้ามี)",
		"สถานะใบงาน",
		"Close note",
		"จำนวนการ Re-open",
		"เหตุผลที่ขอดำเนินเรื่องต่อ",
		"รายละเอียดที่ขอดำเนินเรื่องต่อ",
		"วันที่รับคำขอ"


	];
	var reqCol_new = [
		"หมายเลข Case Task",
		"เรื่อง",
		"สถานะการดำเนินงาน",
		"SLA Due date",
		"กลุ่มที่รับผิดชอบ",
		"Due date",
		"ผู้รับมอบหมาย",
		"หมายเลข Case",
		"ขอรับบริการในนาม",
		"ผู้ติดต่อ",
		"ผู้ขอรับบริการ",
		"ประเภทเลขที่ยืนยันตัวตน",
		"ประเภทเลขที่ยืนยันตัวตน (นิติบุคคล)",
		"เลขที่ยืนยันตัวตน",
		"เลขที่ยืนยันตัวตน (นิติบุคคล)",
		"หมายเลขโทรศัพท์",
		"อีเมล",
		"หมายเลขโทรศัพท์ (สำรอง)",
		"อีเมล (สำรอง)",
		"จังหวัด (ที่อยู่ลูกหนี้)",
		"ชื่อนิติบุคคล",
		"ผู้ให้บริการ",
		"ผลิตภัณฑ์",
		"สถานะบัญชี",
		"โครงการแก้หนี้",
		"เลขที่ทะเบียนรถ",
		"จังหวัดที่จดทะเบียนรถ",
		"แนวทางที่ต้องการเสนอให้ผู้บริการทางการเงินพิจารณา",
		"เหตุผลหรือคำอธิบายประกอบคำขอ",
		"Case ที่เกี่ยวข้อง",
		"Ref Case Task",
		"เลขที่บัตร/ เลขที่สัญญา",
		"วันที่เริ่มติดต่อลูกหนี้",
		"ผลการพิจารณา",
		"แนวทางการช่วยเหลือลูกหนี้",
		"เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้",
		"ภาระหนี้รวม (บาท)",
		"เงินต้น (บาท)",
		"ภาระหนี้ที่ตกลงชำระ (บาท)",
		"จำนวนงวดที่ชำระ (เดือน)",
		"ค่างวดต่อเดือน (บาท)",
		"รายงาน RDT",
		"รายงาน DRD",
		"ข้อมูลที่ต้องการแจ้ง ธปท. เพิ่มเติม (ถ้ามี)",
		"สถานะใบงาน",
		"Close note",
		"จำนวนการ Re-open",
		"เหตุผลที่ขอดำเนินเรื่องต่อ",
		"รายละเอียดที่ขอดำเนินเรื่องต่อ",
		"วันที่รับคำขอ"


	];

	var walkin = ["ผู้ขอรับบริการ",	"เลขที่ยืนยันตัวตน", "เลขที่ยืนยันตัวตน (นิติบุคคล)","หมายเลขโทรศัพท์","อีเมล","ชื่อนิติบุคคล","ผู้ให้บริการ","ผลิตภัณฑ์","สถานะบัญชี","เลขที่ทะเบียนรถ","จังหวัดที่จดทะเบียนรถ", "แนวทางที่ต้องการเสนอให้ผู้บริการทางการเงินพิจารณา", "เลขที่บัตร/ เลขที่สัญญา", "ผลการพิจารณา", "แนวทางการช่วยเหลือลูกหนี้", "เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้", "เงินต้น (บาท)", "วันที่ทำสัญญา","ข้อมูลที่ต้องการแจ้ง ธปท. เพิ่มเติม (ถ้ามี)","วันที่รับคำขอ"];

	var templateVersion = gs.getProperty("fi_template_version");
	var parser = new sn_impex.GlideExcelParser();
	var attachment = new GlideSysAttachment();
	var attachmentStream = attachment.getContentStream(sysId);
	parser.setSource(attachmentStream);
	var list_sheet_name = parser.getSheetNames();
	parser.setSheetName(list_sheet_name[0]); // Select First Sheet Excel
	if (parser.parse()) {
		var headers = parser.getColumnHeaders(); // retrieve the column headers
		if (parser.next()) {
			var header_row2 = Object.values(parser.getRow());
		}

		if (JSON.stringify(headers) == JSON.stringify(reqCol_old) || JSON.stringify(headers) == JSON.stringify(reqCol_new)) {
			gs.info("ไฟล์ excel ถูกต้อง ");
			return 1;
		}
		else if (JSON.stringify(header_row2) == JSON.stringify(walkin)) {
			if(headers[0] != templateVersion){
				current.u_file_error = "เวอร์ชั่นไฟล์ผิดพลาด ปัจจุบันคือ "+templateVersion;
				current.state = "5";
				return false;
			}else{
				gs.info("ไฟล์ excel ถูกต้อง ");
				return 2;
			}
		}
		else {
			current.state = "5";
			// update log error in function
			gs.info("columns ในไฟล์ excel ผิดพลาด!");
			if (headers[0] == "หมายเลข Case Task") {
				logColMissing(reqCol_old, headers, current);
			} else {
				logColMissing(walkin, header_row2, current);
			}

			return false;
		}
	} else {
		current.state = "5"; // Error
		gs.addInfoMessage("ไม่พบ Column ที่ตามรูปแบบที่กำหนด");
		current.u_file_error = "ไม่พบ Column ที่ตามรูปแบบที่กำหนด";
		return false;
	}

}

function checkExcelFile(fileName) {

	var lowercaseFileName = fileName.toLowerCase();
	var fileExtension = lowercaseFileName.split('.').pop();

	if (fileExtension == 'xls' || fileExtension == 'xlsx') {
		return true;
	} else {
		return false;
	}
}

function logColMissing(headersV, excelCol,current) {

	missingColumns = headersV.filter(col => !excelCol.includes(col));

	if (missingColumns.length === headersV.length) {

		gs.addInfoMessage("ไม่พบ Column ที่ตามรูปแบบที่กำหนด");
		current.u_file_error = "ไม่พบ Column ที่ตามรูปแบบที่กำหนด";

	} else if (missingColumns.length > 0) {

		gs.addInfoMessage("ไม่พบ Column " + missingColumns.join(", "));
		current.u_file_error = "ไม่พบ Column " + missingColumns.join(", ");

	}

}