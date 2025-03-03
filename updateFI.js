        var recordsLimit = gs.getProperty('x_baot_debt_sett_0.excelRecordsLimit');
		var countCondition = false;
		if (recordsLimit < 1) countCondition = count < 1;
		else countCondition = count < 1 || count > recordsLimit;

		if (countCondition) {
			var gr = new GlideRecord('sys_attachment');
            gr.addEncodedQuery("table_sys_id=" + source.sys_import_set.data_source.sys_id);
            gr.setLimit(1);
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
            // while loop 
        }