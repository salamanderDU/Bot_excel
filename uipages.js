function onLoad() {

    //const xl = require('excel4node');
    var c = this
    var gamsg = new GlideAjax('global.GlideRecordProcessor');
    gamsg.addParam('sysparm_name', 'getMSG');
    gamsg.getXMLAnswer(function(response) {
        var max = response;
        document.getElementById('max').innerHTML = `<b><u>หมายเหตุ</u></b> ระบบสามารถ Export ได้สูงสุด ` + max + " รายการ";
    });

	var versionMsg = new GlideAjax('global.GlideRecordProcessor');
	versionMsg.addParam('sysparm_name', 'getVersion');
    versionMsg.getXMLAnswer(function(response) {
		    document.getElementById('versionMessage').innerHTML = "โครงการ คุณสู้ เราช่วย กรณีติดต่อผ่านทางผู้ให้บริการ "+response;

	});


    // Fetch the current user's email and set it to the email field
    // var gar = new GlideAjax('global.GlideRecordProcessor');
    // gar.addParam('sysparm_name', 'getFairChoices');
    // gar.getXMLAnswer(function(response) {
    //     var fairchoices = JSON.parse(response);
    //     var fairSelect = document.getElementById('fair');
    //     var fairContainer = document.getElementById('fairContainer');
    //     alert(fairchoices);
    //     // Populate the dropdown with the retrieved values
    //     fairchoices.forEach(function(list) {
    //         var option1 = document.createElement('option');
    //         option1.value = list.value;
    //         option1.text = list.label;
    //         fairSelect.add(1);
    //     });
    // });


    var ga = new GlideAjax('global.GlideRecordProcessor');
    ga.addParam('sysparm_name', 'getSOChoices');
    ga.getXMLAnswer(function(response) {
        var choices = JSON.parse(response);
        var soSelect = document.getElementById('solist');
        var soFieldContainer = document.getElementById('soFieldContainer');

        // Populate the dropdown with the retrieved values
        //alert(choices);
        choices.forEach(function(choice) {
            var option = document.createElement('option');
            option.value = choice.value;
			console.log('option.value:',option.value);
            option.text = choice.label;
            soSelect.add(option);
        });
    });



    var ta = new GlideAjax('global.GlideRecordProcessor');

    ta.addParam('sysparm_name', 'userHasRole');
    //ga.addParam('role', 'admin');
    ta.getXMLAnswer(function(response) {

        var userRole = response;

        if (userRole == 'internal') {
            soFieldContainer.style.display = 'inline';
            //solist.style.display = 'inline';
        }

    });



}

function getChoice() {
    //  alert('test1');
    //  var status = gel('state').value;
    var startDate = gel('date').value;
    var endDate = gel('date1').value;

    var date = document.getElementById('date').value;
    var date1 = document.getElementById('date1').value;

    //  var state = document.getElementById('state').value;
    var solist = document.getElementById('solist').value;
    var fair = document.getElementById('fair').value;
    var type = document.getElementById('exportType').value;
    const selectedItems = $("#glideList").val();
    //var test = document.getElementById('glideList').value;
 //   alert(type);
    // var email = gel('email').value;

    if (!selectedItems || !startDate || !endDate) {
        alert('Please fill out all mandatory fields.');
        return;
    }

    // Change format start date
    var dateRegex = /^\d{1,2}-\d{1,2}-\d{2,4}$/;
    //var dateRegex = /^\d{1,2}\/\d{1,2}\/\d{2,4}$/;



    if (!startDate.match(dateRegex)) {
        alert('Please enter a valid Start date in DD-MM-YYYY format');
        return;
    }
    var startFormat = startDate.split('-');

    var start = '';
    for (var i = startFormat.length; i >= 1 && i <= startFormat.length; i--) {
        start = start + startFormat[i - 1];
        if (i != 1) {
            start = start + "-";
        }
    }
    var start = new Date(start); // get start_date. convert from user's date format


    // Change format end date

    if (!endDate.match(dateRegex)) {
        alert('Please enter a valid End date in DD-MM-YYYY format');
        return;
    }
    var endFormat = endDate.split('-');
    var end = '';
    for (var j = endFormat.length; j >= 1 && j <= endFormat.length; j--) {
        end = end + endFormat[j - 1];
        if (j != 1) {
            end = end + "-";
        }
    }
    var end = new Date(end) //get end_date. convert from user's date format
    console.log(end);
    var currentDate = new Date();

    if (end < start) {
        alert('End date should be more than Start date.');
        return;
    }

    if (start > currentDate) {
        alert('Start date should not on future.');
        return;
    }

    if (end > currentDate) {
        alert('End date should not on future.');
        return;
    }
    var gah = new GlideAjax('global.GlideRecordProcessor');
    gah.addParam('sysparm_name', 'userHasRole');
    gah.getXMLAnswer(function(response) {
        var userRole = response;
        //
        var gar = new GlideAjax('global.GlideRecordProcessor');
        gar.addParam('sysparm_name', 'getRowCountLimit');
        gar.addParam('sysparm_status', selectedItems);
		console.log('selectedItems:',selectedItems);
        gar.addParam('sysparm_date', date);
        gar.addParam('sysparm_date1', date1);
        gar.addParam('sysparm_role', userRole);
        gar.addParam('sysparm_solist', solist);
        gar.addParam('sysparm_fair', fair);
        gar.addParam('sysparm_debt_type', type);
        gar.getXMLAnswer(function(response) {
            alert(response);
            if (response == true || response == "true") {

                var ga = new GlideAjax('global.GlideRecordProcessor');
                ga.addParam('sysparm_name', 'getChoiceRecordsPipe');
                ga.getXMLAnswer(function(response) {
                    var choice = JSON.parse(response);
                    getRole(choice)
                });
            } else if (response == false || response == "false") {
                var gamsg = new GlideAjax('global.GlideRecordProcessor');
                gamsg.addParam('sysparm_name', 'getMSG');
                gamsg.getXMLAnswer(function(response) {
                    alert("ไม่สามารถ Export เกินจำนวน " + response + " รายการ โปรดเลือกช่วงวันใหม่");
                });

            } else {
                alert("ไม่มีข้อมูลในช่วงเวลาที่เลือก กรุณาเลือกช่วงเวลาใหม่");
            }
        });
    });




}

function getRole(choice) {
    var ga = new GlideAjax('global.GlideRecordProcessor');

    ga.addParam('sysparm_name', 'userHasRole');
    //ga.addParam('role', 'admin');
    ga.getXMLAnswer(function(response) {

        var userRole = response;
        //  alert(userRole);
        //document.getElementById('email').value = userEmail;
        getproductChoice(choice, userRole);

    });
}

function getproductChoice(choice, userRole) {


    var ga = new GlideAjax('global.GlideRecordProcessor');

    ga.addParam('sysparm_name', 'getproductRecordsPipe');
    ga.getXMLAnswer(function(response) {



        var productchoice = JSON.parse(response);

        getconsiderationChoice(choice, userRole, productchoice)
    });

}

function getconsiderationChoice(choice, userRole, productchoice) {


    var ga = new GlideAjax('global.GlideRecordProcessor');

    ga.addParam('sysparm_name', 'getconsiderationRecords');
    ga.getXMLAnswer(function(response) {



        var consideration = JSON.parse(response);

        getglidelineChoice(choice, userRole, productchoice, consideration)
    });



}

function getglidelineChoice(choice, userRole, productchoice, consideration) {


    var ga = new GlideAjax('global.GlideRecordProcessor');

    ga.addParam('sysparm_name', 'getguidelinesRecordsPipe');
    ga.getXMLAnswer(function(response) {



        var glideline = JSON.parse(response);

        getreasonChoice(choice, userRole, productchoice, consideration, glideline)
    });



}

function getreasonChoice(choice, userRole, productchoice, consideration, glideline) {


    var ga = new GlideAjax('global.GlideRecordProcessor');

    ga.addParam('sysparm_name', 'getreasonRecordsPipe');
    ga.getXMLAnswer(function(response) {



        var reason = JSON.parse(response);

        getStateChoice(choice, userRole, productchoice, consideration, glideline, reason)
    });



}


function getStateChoice(choice, userRole, productchoice, consideration, glideline, reason) {


    var ga = new GlideAjax('global.GlideRecordProcessor');

    ga.addParam('sysparm_name', 'getstateRecords');
    ga.getXMLAnswer(function(response) {



        var casestate = JSON.parse(response);

        getcancelChoice(choice, userRole, productchoice, consideration, glideline, reason, casestate)
    });



}

function getcancelChoice(choice, userRole, productchoice, consideration, glideline, reason, casestate) {


    var ga = new GlideAjax('global.GlideRecordProcessor');

    ga.addParam('sysparm_name', 'getcancelRecords');
    ga.getXMLAnswer(function(response) {



        var casecancel = JSON.parse(response);

        generateReport(choice, userRole, productchoice, consideration, glideline, reason, casestate, casecancel)
    });



}

function generateReport(choice, userRole, productchoice, consideration, glideline, reason, casestate, casecancel) {

    var date = document.getElementById('date').value;
    var date1 = document.getElementById('date1').value;
    //   var state = document.getElementById('state').value;
    //  var email = document.getElementById('email').value;
    var solist = document.getElementById('solist').value;
    var fair = document.getElementById('fair').value;
    const selectedItems = $("#glideList").val();
    
    var type = document.getElementById('exportType').value;
//	alert('lll'+type)




    var ga = new GlideAjax('global.GlideRecordProcessor');

    ga.addParam('sysparm_name', 'getRecords');
    ga.addParam('sysparm_status', selectedItems);
    ga.addParam('sysparm_date', date);
    ga.addParam('sysparm_date1', date1);
    ga.addParam('sysparm_role', userRole);
    ga.addParam('sysparm_solist', solist);
    ga.addParam('sysparm_fair', fair);

    ga.addParam('sysparm_debt_types', type);
    // alert(solist);
    //  alert(email);
    ga.getXMLAnswer(function(response) {

        // var result = response.responseXML.documentElement.getAttribute("answer");

        var records = JSON.parse(response);
    //    alert(records);

        generateExcel(records, choice, productchoice, consideration, glideline, reason, casestate, casecancel);
    });

    function generateExcel(records, choice, productchoice, consideration, glideline, reason, casestate, casecancel) {
        var choiceRecord = choice;
        var productRecord = productchoice;
        var considerationRecord = consideration;
        var glidelineRecord = glideline;
        var reasonRecord = reason;
        var casestatus = casestate;
        var cancel = casecancel;

        // alert(JSON.stringify(choiceRecord));
        alert(JSON.stringify(records));
        var workbook = new ExcelJS.Workbook();
        //xl.workbook
        var worksheet = workbook.addWorksheet('File1');
        const sheet = workbook.addWorksheet('File2');
        const sheet2 = workbook.addWorksheet('File3');
        const sheet3 = workbook.addWorksheet('File4');
        const sheet4 = workbook.addWorksheet('File5');
        const sheet5 = workbook.addWorksheet('File6');
        const sheet6 = workbook.addWorksheet('File7');
        const sheet7 = workbook.addWorksheet('File8');


        worksheet.addRow(['หมายเลข Case Task', 'เรื่อง', 'สถานะการดำเนินงาน', 'SLA Due date', 'กลุ่มที่รับผิดชอบ', 'Due date', 'ผู้รับมอบหมาย', 'หมายเลข Case', 'ขอรับบริการในนาม', 'ผู้ติดต่อ', 'ผู้ขอรับบริการ', 'ประเภทเลขที่ยืนยันตัวตน', 'ประเภทเลขที่ยืนยันตัวตน (นิติบุคคล)', 'เลขที่ยืนยันตัวตน', 'เลขที่ยืนยันตัวตน (นิติบุคคล)', 'หมายเลขโทรศัพท์', 'อีเมล', 'หมายเลขโทรศัพท์ (สำรอง)', 'อีเมล (สำรอง)', 'จังหวัด (ที่อยู่ลูกหนี้)', 'ชื่อนิติบุคคล', 'ผู้ให้บริการ', 'ผลิตภัณฑ์', 'สถานะบัญชี', 'โครงการแก้หนี้', 'เลขที่ทะเบียนรถ', 'จังหวัดที่จดทะเบียนรถ', 'แนวทางที่ต้องการเสนอให้ผู้บริการทางการเงินพิจารณา', 'เหตุผลหรือคำอธิบายประกอบคำขอ', 'Case ที่เกี่ยวข้อง', 'Ref Case Task', 'เลขที่บัตร/ เลขที่สัญญา', 'วันที่เริ่มติดต่อลูกหนี้', 'ผลการพิจารณา', 'แนวทางการช่วยเหลือลูกหนี้', 'เหตุผลที่ไม่สามารถช่วยเหลือลูกหนี้ได้', 'ภาระหนี้รวม (บาท)', 'เงินต้น (บาท)', 'ภาระหนี้ที่ตกลงชำระ (บาท)', 'จำนวนงวดที่ชำระ (เดือน)', 'ค่างวดต่อเดือน (บาท)', 'รายงาน RDT', 'วันที่ทำสัญญา', 'ข้อมูลที่ต้องการแจ้ง ธปท. เพิ่มเติม (ถ้ามี)', 'สถานะใบงาน', 'Close note', 'จำนวนการ Re-open', 'เหตุผลที่ขอดำเนินเรื่องต่อ', 'รายละเอียดที่ขอดำเนินเรื่องต่อ', 'วันที่รับคำขอ']);


        choiceRecord.forEach(function(record) {
            sheet.addRow([record.name]);
        });

        sheet.state = 'hidden';
        // sheet.addRow(['_']);
        records.forEach(function(record) {


            if (record.u_number_car == null || record.u_number_car == 'null' || record.u_number_car == '') {

                record.u_number_car = '';
            } else {
                if (record.u_number_car.startsWith("=")) {

                    record.u_number_car = `'` + record.u_number_car;
                }
            }

            if (record.u_reason == null || record.u_reason == 'null' || record.u_reason == '') {

                record.u_reason = '';
            } else {
                if (record.u_reason.startsWith("=")) {
                    record.u_reason = `'` + record.u_reason;
                }
            }

            if (record.u_number_contact == null || record.u_number_contact == 'null' || record.u_number_contact == '') {

                record.u_number_contact = '';
            } else {
                if (record.u_number_contact.startsWith("=")) {
                    record.u_number_contact = `'` + record.u_number_contact;
                }
            }

            if (record.u_number_juristic == null || record.u_number_juristic == 'null' || record.u_number_juristic == '') {

                record.u_number_juristic = '';
            } else {
                if (record.u_number_juristic.startsWith("=")) {
                    record.u_number_juristic = `'` + record.u_number_juristic;
                }
            }


            if (record.u_juristic_name == null || record.u_juristic_name == 'null' || record.u_juristic_name == '') {

                record.u_juristic_name = '';
            } else {
                if (record.u_juristic_name.startsWith("=")) {
                    record.u_juristic_name = `'` + record.u_juristic_name;
                }
            }


            worksheet.addRow([record.number, record.short_description, record.u_work_state, record.u_sla_due_date, record.assignment_group, record.due_date, record.assigned_to, record.u_display_number, record.u_receive_service, record.u_exact_consumer, record.u_consumer, record.u_type_iden, record.u_type_juristic, record.u_iden_number, record.u_number_juristic, record.u_mobile, record.u_mail, record.u_secondary_phone, record.u_secondary_email, record.u_province, record.u_juristic_name, record.u_provider, record.u_product, record.u_state_debt, record.u_debt_project, record.u_number_car, record.u_province_car, record.u_offer_provider, record.u_reason, record.u_major_case, '', record.u_number_contact, record.u_contact_debt, record.u_resolution_code, record.u_glideline_debt, record.u_reason_unable_help, record.u_debt_burden, record.u_principle, record.u_debt_confirm, record.u_installment, record.u_amount_month, record.u_rdt, record.u_drd, record.u_notify_debtor, record.state, '', record.u_re_open_count, record.u_reason_reopen_choice, record.u_reason_reopen, record.opened_at]);
        });


        sheet2.state = 'hidden';
        // sheet2.addRow(['m']);
        productRecord.forEach(function(record) {
            sheet2.addRow([record.name]);
        });
        sheet3.state = 'hidden';
        //sheet3.addRow(['m']);
        considerationRecord.forEach(function(record) {
            sheet3.addRow([record.name]);
        });

        sheet4.state = 'hidden';
        //       sheet4.addRow(['m']);
        glidelineRecord.forEach(function(record) {
            sheet4.addRow([record.name]);
        });

        sheet5.state = 'hidden';
        //sheet5.addRow(['m']);
        reasonRecord.forEach(function(record) {
            sheet5.addRow([record.name]);
        });


        sheet6.state = 'hidden';
        //sheet5.addRow(['m']);
        casestatus.forEach(function(record) {
            sheet6.addRow([record.name]);
        });

        sheet7.state = 'hidden';
        //sheet5.addRow(['m']);
        cancel.forEach(function(record) {
            sheet7.addRow([record.name]);
        });

        // Set cell colors
        worksheet.getCell('A1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: 'DE9D9B'
            } // red
        };
        worksheet.getCell('B1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: 'DE9D9B'
            } // red
        };
        worksheet.getCell('C1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: 'DE9D9B'
            } // red
        };

        worksheet.getCell('D1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: 'DE9D9B'
            } // red
        };
        worksheet.getCell('E1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: 'DE9D9B'
            } // red
        };
        worksheet.getCell('F1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: 'DE9D9B'
            } // red
        };
        worksheet.getCell('G1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: '9DC384'
            } // Green
        };
        worksheet.getCell('H1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: 'DE9D9B'
            } // red
        };
        worksheet.getCell('I1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: 'DE9D9B'
            } // red
        };
        worksheet.getCell('J1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: 'DE9D9B'
            } // red
        };
        worksheet.getCell('K1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: 'DE9D9B'
            } // red
        };
        worksheet.getCell('L1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: 'DE9D9B'
            } // red
        };
        worksheet.getCell('M1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: 'DE9D9B'
            } // red
        };
        worksheet.getCell('N1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: 'DE9D9B'
            } // red
        };
        worksheet.getCell('O1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: 'DE9D9B'
            } // red
        };
        worksheet.getCell('P1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: 'DE9D9B'
            } // red
        };
        worksheet.getCell('Q1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: 'DE9D9B'
            } // red
        };
        worksheet.getCell('R1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: 'DE9D9B'
            } // red
        };
        worksheet.getCell('S1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: 'DE9D9B'
            } // red
        };
        worksheet.getCell('T1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: 'DE9D9B'
            } // red
        };
        worksheet.getCell('U1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: 'DE9D9B'
            } // red
        };
        worksheet.getCell('V1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: 'DE9D9B'
            } // red
        };
        worksheet.getCell('W1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: '9DC384'
            } // green
        };
        worksheet.getCell('X1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: '9DC384'
            } // green
        };
        worksheet.getCell('Y1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: 'DE9D9B'
            } // red
        };
        worksheet.getCell('Z1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: '9DC384'
            } // green
        };

        worksheet.getCell('AA1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: '9DC384'
            } // green
        };

        worksheet.getCell('AB1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: 'DE9D9B'
            } // red
        };

        worksheet.getCell('AC1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: 'DE9D9B'
            } // red
        };

        worksheet.getCell('AD1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: '9DC384'
            } // green
        };
        worksheet.getCell('AE1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: '9DC384'
            } // green
        };
        worksheet.getCell('AF1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: '9DC384'
            } // green
        };

        worksheet.getCell('AG1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: '9DC384'
            } // green
        };

        worksheet.getCell('AH1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: '9DC384'
            } // green
        };

        worksheet.getCell('AI1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: '9DC384'
            } // green
        };

        worksheet.getCell('AJ1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: '9DC384'
            } // green
        };
        worksheet.getCell('AK1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: '9DC384'
            } // green
        };
        worksheet.getCell('AL1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: '9DC384'
            } // green
        };
        worksheet.getCell('AM1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: '9DC384'
            } // green
        };
        worksheet.getCell('AN1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: '9DC384'
            } // green
        };
        worksheet.getCell('AO1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: '9DC384'
            } // green
        };
        worksheet.getCell('AP1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: '9DC384'
            } // green
        };

        worksheet.getCell('AQ1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: '9DC384'
            } // green
        };

        worksheet.getCell('AR1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: '9DC384'
            } // green
        };
        worksheet.getCell('AS1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: '9DC384'
            } // green
        };
        worksheet.getCell('AT1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: '9DC384'
            } // green
        };
        worksheet.getCell('AU1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: 'DE9D9B'
            } // red
        };
        worksheet.getCell('AV1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: 'DE9D9B'
            } // red
        };
        worksheet.getCell('AW1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: 'DE9D9B'
            } // red
        };
        worksheet.getCell('AX1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                argb: 'DE9D9B'
            } // red
        };



        worksheet.getColumn("X").eachCell({
            includeEmpty: true
        }, (cell, rowNumber) => {
            cell.dataValidation = {

                type: 'list',
                allowBlank: true,

                showErrorMessage: true,

                formulae: ['File2!$A$1:$A$500'],
                // errorStyle: 'error',
                // error: 'The value Valid',
            };
        });

        worksheet.getColumn("W").eachCell({
            includeEmpty: true
        }, function(cell, rowNumber) {
            cell.dataValidation = {
                type: 'list',
                allowBlank: true,
                showErrorMessage: true,
                formulae: ['File3!$A$1:$A$500'],
                // errorStyle: 'error',
                // error: 'The value Valid',
            };
        });

        worksheet.getColumn("AH").eachCell({
            includeEmpty: true
        }, function(cell, rowNumber) {
            cell.dataValidation = {
                type: 'list',
                allowBlank: true,
                showErrorMessage: true,
                formulae: ['File4!$A$1:$A$500'],
                // errorStyle: 'error',
                // error: 'The value Valid',
            };
        });

        worksheet.getColumn("AI").eachCell({
            includeEmpty: true
        }, function(cell, rowNumber) {
            cell.dataValidation = {
                type: 'list',
                allowBlank: true,
                showErrorMessage: true,
                formulae: ['File5!$A$1:$A$500'],
                // errorStyle: 'error',
                // error: 'The value Valid',
            };
        });

        worksheet.getColumn("AJ").eachCell({
            includeEmpty: true
        }, function(cell, rowNumber) {
            cell.dataValidation = {
                type: 'list',
                allowBlank: true,
                showErrorMessage: true,
                formulae: ['File6!$A$1:$A$500'],

            };
        });

        worksheet.getColumn("AS").eachCell({
            includeEmpty: true
        }, function(cell, rowNumber) {
            cell.dataValidation = {
                type: 'list',
                allowBlank: true,
                showErrorMessage: true,
                formulae: ['File7!$A$1:$A$500'],
            };
        });

        var password = 'test123'

        // workbook.xlsx.writeBuffer({
        //     password: test123
        // }).then((buffer) => {
        workbook.xlsx.writeBuffer().then(function(buffer) {



            var blob = new Blob([buffer], {
                type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            });



            const base64Data = buffer.toString('base64');
            var url = URL.createObjectURL(blob);

            var g_form = new GlideForm();



            var ga = new GlideAjax('global.ExcelProcessor');
            ga.addParam('sysparm_name', 'uploadFile');
            ga.addParam('sysparm_content', buffer);
            // ga.addParam('sysparm_password', 'kanom313326');
            ga.getXML(function(response) {
                g_form.clearMessages();
                var answer = response.responseXML.documentElement.getAttribute("answer");
                if (answer.startsWith("Export File Success")) {
                    //workbook.setWorkbookPassword("hide-passwd", HashAlgorithm.sha512);
                    g_form.addInfoMessage(new GwtMessage().getMessage("{0}", answer));
                } else {
                    g_form.addInfoMessage(new GwtMessage().getMessage("{0}", answer));
                }

            });

            var now = new Date();
            var year = now.getFullYear();
            var month = ('0' + (now.getMonth() + 1)).slice(-2);
            var day = ('0' + now.getDate()).slice(-2);
            var hours = ('0' + now.getHours()).slice(-2);
            var minutes = ('0' + now.getMinutes()).slice(-2);
            var seconds = ('0' + now.getSeconds()).slice(-2);

            // Format the date and time
            // + '_' + hours + minutes + seconds
            var formattedDateTime = day + month + year;
            var link = document.createElement('a');
            link.href = url;
            link.download = '[Debt Settlement]Bulk_Upload_' + formattedDateTime + '.xlsx';
            link.click();

        });


    }

}
