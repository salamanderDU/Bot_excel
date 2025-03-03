var fi = [0, 1, 3, 4, 5, 9, 17, 18, 22, 24, 25, 26, 27];
// length of the array = 13
// console.log(fi.length); 

var newHeader = ["ผู้ขอรับบริการ", "เลขที่ยืนยันตัวตน", "เลขที่ยืนยันตัวตน (นิติบุคคล)", "หมายเลขโทรศัพท์", "อีเมล", "ชื่อนิติบุคคล", "ผู้ให้บริการ", "ผลิตภัณฑ์", "สถานะบัญชี", "เลขที่ทะเบียนรถ", "จังหวัดที่จดทะเบียนรถ", "แนวทางที่ต้องการเสนอให้ผู้บริการทางการเงินพิจารณา", "เลขที่บัตร/ เลขที่สัญญา", "ผลการพิจารณา", "แนวทางการช่วยเหลือลูกหนี้", "เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้", "เงินต้น (บาท)", "วันที่ทำสัญญา", "ข้อมูลที่ต้องการแจ้ง ธปท. เพิ่มเติม (ถ้ามี)", "วันที่รับคำขอ"];

var walkin = ["ขอรับบริการในนาม", "ผู้ติดต่อ", "ผู้ขอรับบริการ", "ประเภทเลขที่ยืนยันตัวตน", "ประเภทเลขที่ยืนยันตัวตน (นิติบุคคล)", "เลขที่ยืนยันตัวตน", "เลขที่ยืนยันตัวตน (นิติบุคคล)", "หมายเลขโทรศัพท์", "อีเมล", "จังหวัด (ที่อยู่ลูกหนี้)", "ชื่อนิติบุคคล", "ผู้ให้บริการ", "ผลิตภัณฑ์", "สถานะบัญชี", "เลขที่ทะเบียนรถ", "จังหวัดที่จดทะเบียนรถ", "แนวทางที่ต้องการเสนอให้ผู้บริการทางการเงินพิจารณา", "เลขที่บัตร/ เลขที่สัญญา", "วันที่เริ่มติดต่อลูกหนี้", "ผลการพิจารณา", "แนวทางการช่วยเหลือลูกหนี้", "เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้", "ภาระหนี้รวม (บาท)", "เงินต้น (บาท)", "ภาระหนี้ที่ตกลงชำระ (บาท)", "จำนวนงวดที่ชำระ (เดือน)", "ค่างวดต่อเดือน (บาท)", "รายงาน RDT", "วันที่ทำสัญญา", "ข้อมูลที่ต้องการแจ้ง ธปท. เพิ่มเติม (ถ้ามี)", "วันที่รับคำขอ"];
// what in wlakin is not in newHeader
var diff = walkin.filter(x => !newHeader.includes(x));
// console.log(diff);

var diff2 = ['ขอรับบริการในนาม', 'ผู้ติดต่อ', 'ประเภทเลขที่ยืนยันตัวตน', 'ประเภทเลขที่ยืนยันตัวตน (นิติบุคคล)', 'จังหวัด (ที่อยู่ลูกหนี้)', 'วันที่เริ่มติดต่อลูกหนี้', 'ภาระหนี้รวม (บาท)', 'ภาระหนี้ที่ตกลงชำระ (บาท)', 'จำนวนงวดที่ชำระ (เดือน)', 'ค่างวดต่อเดือน (บาท)', 'รายงาน RDT'];

// what in wlakin is not in newHeader index version statring from 0
var diffIndex = diff2.map(x => walkin.indexOf(x));
// console.log(diffIndex);

var codeDiff = [0, 1, 3, 4, 9, 18, 22, 24, 25, 26, 27];
var diff2 = ['ขอรับบริการในนาม', 'ผู้ติดต่อ', 'ประเภทเลขที่ยืนยันตัวตน', 'ประเภทเลขที่ยืนยันตัวตน (นิติบุคคล)', 'จังหวัด (ที่อยู่ลูกหนี้)', 'วันที่เริ่มติดต่อลูกหนี้', 'ภาระหนี้รวม (บาท)', 'ภาระหนี้ที่ตกลงชำระ (บาท)', 'จำนวนงวดที่ชำระ (เดือน)', 'ค่างวดต่อเดือน (บาท)', 'รายงาน RDT'];


var casetaskGrFiled_FI = ["u_bulk_upload", "short_description", "u_show_on_ticket_list_page", "u_walk_service_requester", "u_walkin_identification_number", "u_walkin_receive_service", "u_walkin_phone", "u_walkin_email", "u_walkin_juristic_name", "u_identification_number_juristic", "u_provider", "assignment_group", "sys_created_by", "assigned_to", "u_product", "u_state_debt", "u_debt_project", "u_number_car", "u_province_car", "u_offer_provider", "u_number_contact", "u_glideline_debt", "u_resolution_code", "state", "u_work_state", "close_notes", "u_reason_unable_help", "u_principle", "u_drd", "opened_at"];
var casetaskGrFiled_FI_onlyin_template = ["u_walk_service_requester", "u_walkin_identification_number", "u_walkin_phone", "u_walkin_email", "u_walkin_juristic_name", "u_identification_number_juristic", "u_provider", "u_product", "u_state_debt", "u_number_car", "u_province_car", "u_offer_provider", "u_number_contact", "u_glideline_debt", "u_resolution_code", "close_notes", "u_reason_unable_help", "u_principle", "u_drd", "opened_at"];
var nameField = ["number", "short_description", "u_consumer", "due_date", "u_sla_due_date", "assignment_group", "assigned_to", "u_receive_service", "u_exact_consumer", "u_display_number", "u_type_juristic", "u_type_iden", "u_iden_number", "u_number_juristic", "u_juristic_name", "u_mobile", "u_secondary_phone", "u_mail", "u_secondary_email", "u_province", "u_provider", "u_debt_project", "u_offer_provider", "u_reason", "u_major_case", "u_debt_detail", "u_reason_reopen_choice", "u_reason_reopen", "u_work_state", "u_identidifier_type", "u_phone_numer", "state", "u_number_car", "u_province_car", "u_number_contact", "u_contact_debt", "u_resolution_code", "u_product", "u_state_debt", "u_glideline_debt", "u_reason_unable_help", "u_debt_burden", "u_principle", "u_debt_confirm", "u_installment", "u_amount_month", "u_drd", "u_rdt", "u_notify_debtor", "closed_at", "u_re_open_count", "opened_at"];
// console.log(newHeader.length);

var diffreport = nameField.filter(x => !casetaskGrFiled_FI_onlyin_template.includes(x));
var diffreport2 = casetaskGrFiled_FI_onlyin_template.filter(x => !nameField.includes(x));
// console.log(diffreport2);
// console.log(diffreport2.length);

var excelheader = ['หมายเลข Case Task', 'เรื่อง', 'สถานะการดำเนินงาน', 'SLA Due date', 'กลุ่มที่รับผิดชอบ', 'Due date', 'ผู้รับมอบหมาย', 'หมายเลข Case', 'ขอรับบริการในนาม', 'ผู้ติดต่อ', 'ผู้ขอรับบริการ', 'ประเภทเลขที่ยืนยันตัวตน', 'ประเภทเลขที่ยืนยันตัวตน (นิติบุคคล)', 'เลขที่ยืนยันตัวตน', 'เลขที่ยืนยันตัวตน (นิติบุคคล)', 'หมายเลขโทรศัพท์', 'อีเมล', 'หมายเลขโทรศัพท์ (สำรอง)', 'อีเมล (สำรอง)', 'จังหวัด (ที่อยู่ลูกหนี้)', 'ชื่อนิติบุคคล', 'ผู้ให้บริการ', 'ผลิตภัณฑ์', 'สถานะบัญชี', 'โครงการแก้หนี้', 'เลขที่ทะเบียนรถ', 'จังหวัดที่จดทะเบียนรถ', 'แนวทางที่ต้องการเสนอให้ผู้บริการทางการเงินพิจารณา', 'เหตุผลหรือคำอธิบายประกอบคำขอ', 'Case ที่เกี่ยวข้อง', 'Ref Case Task', 'เลขที่บัตร/ เลขที่สัญญา', 'วันที่เริ่มติดต่อลูกหนี้', 'ผลการพิจารณา', 'แนวทางการช่วยเหลือลูกหนี้', 'เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้', 'ภาระหนี้รวม (บาท)', 'เงินต้น (บาท)', 'ภาระหนี้ที่ตกลงชำระ (บาท)', 'จำนวนงวดที่ชำระ (เดือน)', 'ค่างวดต่อเดือน (บาท)', 'รายงาน RDT', 'วันที่ทำสัญญา', 'ข้อมูลที่ต้องการแจ้ง ธปท. เพิ่มเติม (ถ้ามี)', 'สถานะใบงาน', 'Close note', 'จำนวนการ Re-open', 'เหตุผลที่ขอดำเนินเรื่องต่อ', 'รายละเอียดที่ขอดำเนินเรื่องต่อ', 'วันที่รับคำขอ'];
var recordheader = ["number", "short_description", "u_work_state", "u_sla_due_date", "assignment_group", "due_date", "assigned_to", "u_display_number", "u_receive_service", "u_exact_consumer", "u_consumer", "u_type_iden", "u_type_juristic", "u_iden_number", "u_number_juristic", "u_mobile", "u_mail", "u_secondary_phone", "u_secondary_email", "u_province", "u_juristic_name", "u_provider", "u_product", "u_state_debt", "u_debt_project", "u_number_car", "u_province_car", "u_offer_provider", "u_reason", "u_major_case", "", "u_number_contact", "u_contact_debt", "u_resolution_code", "u_glideline_debt", "u_reason_unable_help", "u_debt_burden", "u_principle", "u_debt_confirm", "u_installment", "u_amount_month", "u_rdt", "u_drd", "u_notify_debtor", "state", "", "u_re_open_count", "u_reason_reopen_choice", "u_reason_reopen", "opened_at"];

// map excelheader and recordheader and make it key value pair
var mapCell = excelheader.map((key, index) => ({ [key]: recordheader[index] })); // อันนี้ของตัว bot
console.log(JSON.stringify(mapCell, null, 2));

// map newHeader and casetaskGrFiled_FI_onlyin_template and make it key value pair
// อันนี้ของตัว FI ว่าหัวนี้เอาฟิลไหน
var mapCell3 = newHeader.map((key, index) => ({ [key]: casetaskGrFiled_FI_onlyin_template[index] }));
// console.log(JSON.stringify(mapCell3, null, 2));
// console.log(newHeader.length + " " + casetaskGrFiled_FI_onlyin_template.length);

var FI_map = [
    {
        "ผู้ขอรับบริการ": "u_walk_service_requester"
    },
    {
        "เลขที่ยืนยันตัวตน": "u_walkin_identification_number"
    },
    {
        "เลขที่ยืนยันตัวตน (นิติบุคคล)": "u_identification_number_juristic" 
    },
    {
        "หมายเลขโทรศัพท์": "u_walkin_phone" 
    },
    {
        "อีเมล": "u_walkin_email" 
    },
    {
        "ชื่อนิติบุคคล": "u_walkin_juristic_name" 
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
        "เลขที่ทะเบียนรถ": "u_number_car"
    },
    {
        "จังหวัดที่จดทะเบียนรถ": "u_province_car"
    },
    {
        "แนวทางที่ต้องการเสนอให้ผู้บริการทางการเงินพิจารณา": "u_offer_provider"
    },
    {
        "เลขที่บัตร/ เลขที่สัญญา": "u_number_contact"
    },
    {
        "ผลการพิจารณา": "u_glideline_debt"
    },
    {
        "แนวทางการช่วยเหลือลูกหนี้": "u_resolution_code"
    },
    {
        "เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้": "close_notes"
    },
    {
        "เงินต้น (บาท)": "u_reason_unable_help"
    },
    {
        "วันที่ทำสัญญา": "u_principle"
    },
    {
        "ข้อมูลที่ต้องการแจ้ง ธปท. เพิ่มเติม (ถ้ามี)": "u_drd"
    },
    {
        "วันที่รับคำขอ": "opened_at"
    }
] // อันนี้ของตัว FI

// find diff between FI_map key and mapCell key
// อันนี้เป็นของ export กับ insert ของ FI จะได้เอามาดูว่าอะไรบ้างที่เรา insert ไป แล้วไม่ออกจากที่ export
var diffFI = mapCell.filter(x => !FI_map.map(y => Object.keys(y)[0]).includes(Object.keys(x)[0]));
// console.log(JSON.stringify(diffFI, null, 2));




