var notcal = ["u_walk_service_requester", "u_walkin_identification_number", "u_walkin_phone", "u_walkin_email", "u_walkin_juristic_name", "u_identification_number_juristic", "u_provider", "u_product", "u_state_debt", "u_offer_provider", "u_number_contact", "u_glideline_debt", "u_resolution_code", "close_notes", "u_reason_unable_help", "u_principle", "u_drd", "opened_at",'u_number_car', 'u_province_car'];
var casetaskGrFiled_FI_onlyin_template = ["u_walk_service_requester", "u_walkin_identification_number", "u_walkin_phone", "u_walkin_email", "u_walkin_juristic_name", "u_identification_number_juristic", "u_provider", "u_product", "u_state_debt", "u_debt_project", "u_number_car", "u_province_car", "u_offer_provider", "u_number_contact", "u_glideline_debt", "u_resolution_code", "close_notes", "u_reason_unable_help", "u_principle", "u_drd", "opened_at"];

// find diff between 2 array
var diff = casetaskGrFiled_FI_onlyin_template.filter(x => !notcal.includes(x));
console.log(diff);

