
function walkinFIbulkUpload(queryTable, query, datee, fairquery, external) {
	var gr = new GlideRecord(queryTable);
	gr.addEncodedQuery(query + datee + fairquery + external);
	gr.query();

	var records = [];
	while (gr.next()) {
		var recordObj = {};
		recordObj.number = gr.getValue('number');
		recordObj.short_description = gr.getValue('short_description');
		recordObj.u_consumer = gr.getDisplayValue('u_name_consumer');
		recordObj.due_date = gr.getDisplayValue('due_date');
		recordObj.u_sla_due_date = gr.getDisplayValue('u_sla_due_date');
		recordObj.assignment_group = gr.getDisplayValue('assignment_group');
		recordObj.assigned_to = gr.getDisplayValue('assigned_to.email');
		// recordObj.u_receive_service = gr.getDisplayValue('u_parent_case.request_by');
		recordObj.u_receive_service = gr.u_receive_service.getDisplayValue();
		recordObj.u_exact_consumer = gr.getDisplayValue('u_name_exact_consumer');
		recordObj.u_display_number = gr.getDisplayValue('u_display_number');
		//recordObj.u_type_juristic = gr.getDisplayValue('u_parent_case.juristic_type');
		recordObj.u_type_juristic = gr.u_type_juristic.getDisplayValue();
		recordObj.u_type_iden = gr.getDisplayValue('u_type_iden');
		recordObj.u_iden_number = gr.getDisplayValue('u_iden_number');
		// recordObj.u_number_juristic = gr.u_parent_case.juristic_number;
		recordObj.u_number_juristic = gr.u_number_juristic.getDisplayValue();
		//  recordObj.u_juristic_name = gr.getDisplayValue('u_parent_case.juristic_name');
		recordObj.u_juristic_name = gr.u_juristic_name.getDisplayValue();
		recordObj.u_mobile = gr.getDisplayValue('u_mobile');
		//   recordObj.u_secondary_phone = gr.getDisplayValue('u_secondary_phone');
		recordObj.u_secondary_phone = gr.u_secondary_phone.getDisplayValue();
		recordObj.u_mail = gr.getDisplayValue('u_mail');
		recordObj.u_secondary_email = gr.getDisplayValue('u_secondary_email');
		recordObj.u_province = gr.getDisplayValue('u_province');
		recordObj.u_provider = gr.getDisplayValue('u_provider');
		recordObj.u_debt_project = gr.getDisplayValue('u_debt_project');
		recordObj.u_offer_provider = gr.getDisplayValue('u_offer_provider');
		recordObj.u_reason = gr.getDisplayValue('u_reason');
		recordObj.u_major_case = gr.getDisplayValue('u_major_case');
		recordObj.u_debt_detail = gr.getDisplayValue('u_debt_detail');
		recordObj.u_reason_reopen_choice = gr.getDisplayValue('u_reason_reopen_choice');
		recordObj.u_reason_reopen = gr.getDisplayValue('u_reason_reopen');
		recordObj.u_work_state = gr.getValue('u_work_state');
		recordObj.u_identidifier_type = gr.u_iden_number.getDisplayValue();
		recordObj.u_phone_numer = gr.u_mobile.getDisplayValue();
		//   recordObj.u_product = gr.getDisplayValue('u_product');
		//  recordObj.u_state_debt = gr.getDisplayValue('u_state_debt');
		recordObj.state = gr.getDisplayValue('state');
		recordObj.u_number_car = gr.getValue('u_number_car');
		recordObj.u_province_car = gr.getDisplayValue('u_province_car');
		recordObj.u_number_contact = gr.getValue('u_number_contact');
		recordObj.u_contact_debt = gr.getDisplayValue('u_contact_debt');
		recordObj.u_resolution_code = gr.getDisplayValue('u_resolution_code');
		//  recordObj.u_glideline_debt = gr.getDisplayValue('u_glideline_debt');
		//  recordObj.u_reason_unable_help = gr.getDisplayValue('u_reason_unable_help');

		var grXBDS_0PTF = new GlideRecord('x_baot_debt_sett_0_product_to_fair');
		grXBDS_0PTF.addEncodedQuery("x_baot_debt_sett_0_debt_fair!=bd12cc231b5dbd1080bb844ee54bcb5c^ORx_baot_debt_sett_0_debt_fair=NULL^u_m2m_product_to_application.sys_id=" + gr.getValue('u_product'));
		grXBDS_0PTF.query();

		if (grXBDS_0PTF.next()) {
			// recordObj.u_product = grXBDS_0PTF.getDisplayValue("u_m2m_product_to_application") + " | " + grXBDS_0PTF.getDisplayValue("x_baot_debt_sett_0_debt_fair");
			recordObj.u_product = grXBDS_0PTF.getDisplayValue("u_m2m_product_to_application") + " | " + gr.getDisplayValue('u_debt_project');
		}

		var grXBDS_0DTFStateDept = new GlideRecord('x_baot_debt_sett_0_debt_to_fair');
		grXBDS_0DTFStateDept.addEncodedQuery("x_baot_debt_sett_0_debt_fair!=a86920c51b2cb1106c24dcace54bcb8a^ORx_baot_debt_sett_0_debt_fair=NULL^x_baot_debt_sett_0_debt_fair!=6904000d1bacf11080bb844ee54bcb95^ORx_baot_debt_sett_0_debt_fair=NULL^x_baot_debt_sett_0_debt_fair!=bd12cc231b5dbd1080bb844ee54bcb5c^ORx_baot_debt_sett_0_debt_fair=NULL^x_baot_debt_sett_0_debt_fair!=cd924c631b5dbd1080bb844ee54bcbbc^ORx_baot_debt_sett_0_debt_fair=NULL^x_baot_debt_sett_0_debt_fair!=dd944f0c1b34b91080bb844ee54bcbf9^ORx_baot_debt_sett_0_debt_fair=NULL^u_m2m_debt_status_to_application.sys_id=" + gr.getValue("u_state_debt"));
		grXBDS_0DTFStateDept.query();

		if (grXBDS_0DTFStateDept.next()) {
			// recordObj.u_state_debt = grXBDS_0DTFStateDept.getDisplayValue("u_m2m_debt_status_to_application") + " | " + grXBDS_0DTFStateDept.getDisplayValue("x_baot_debt_sett_0_debt_fair");
			recordObj.u_state_debt = grXBDS_0DTFStateDept.getDisplayValue("u_m2m_debt_status_to_application") + " | " + gr.getDisplayValue('u_debt_project');
		}

		recordObj.state = gr.getDisplayValue('state');
		recordObj.u_number_car = gr.getValue('u_number_car');
		recordObj.u_province_car = gr.getDisplayValue('u_province_car');
		recordObj.u_number_contact = gr.getValue('u_number_contact');
		recordObj.u_contact_debt = gr.getDisplayValue('u_contact_debt');
		recordObj.u_resolution_code = gr.getDisplayValue('u_resolution_code');

		var grUGD = new GlideRecord('u_glideline_debtor');
		grUGD.addEncodedQuery("sys_id=" + gr.getValue("u_glideline_debt"));
		grUGD.query();

		if (grUGD.next()) {
			recordObj.u_glideline_debt = grUGD.getDisplayValue('u_name') + " | " + grUGD.getDisplayValue('u_debt_fair');
		}

		var grUUH = new GlideRecord('u_unable_help');
		grUUH.addEncodedQuery("sys_id=" + gr.getValue("u_reason_unable_help"));
		grUUH.query();

		if (grUUH.next()) {
			recordObj.u_reason_unable_help = grUUH.getDisplayValue('u_name') + " | " + grUUH.getDisplayValue('u_debt_fair');
		}

		recordObj.u_debt_burden = gr.getDisplayValue('u_debt_burden');
		recordObj.u_principle = gr.getDisplayValue('u_principle');
		recordObj.u_debt_confirm = gr.getDisplayValue('u_debt_confirm');
		recordObj.u_installment = gr.getDisplayValue('u_installment');
		recordObj.u_amount_month = gr.getDisplayValue('u_amount_month');
		recordObj.u_drd = gr.getDisplayValue('u_drd');
		recordObj.u_rdt = gr.getDisplayValue('u_rdt');
		recordObj.u_notify_debtor = gr.getDisplayValue('u_notify_debtor');
		recordObj.closed_at = gr.getDisplayValue('closed_at');
		recordObj.u_re_open_count = gr.getDisplayValue('u_re_open_count');
		recordObj.opened_at = gr.getDisplayValue('opened_at');
		records.push(recordObj);
	}
	return JSON.stringify(records);
}