function check_money_valid_new(amount) {
	const regex = /^(฿)?(\d{1,3}(,\d{3})*|\d+)(\.\d+)?$/;
	return regex.test(amount);
}


function check_money_valid_old(amount) {
	const regex = /^(฿)?(\d{1,3}(,\d{3})*|\d+)(\.\d{1,2})?$/;
	return regex.test(amount);
}

function set_money_object(money_input) {
	var x = money_input.replace("฿", "");
	var y = x.replace(",", "");
	return "THB;" + y;
}

// Test case
console.log(check_money_valid_new('฿9,999.99')); // true
console.log(set_money_object('฿9,999.99'));
console.log(check_money_valid_old('฿9,999.99')); // true
console.log(set_money_object('฿9,999.99'));

console.log(check_money_valid_new('฿9,999.999')); // true
console.log(set_money_object('฿9,999.999'));
console.log(check_money_valid_old('฿9,999.999')); // false
console.log(set_money_object('฿9,999.999'));

console.log(check_money_valid_new('฿9,999.9')); // true
console.log(set_money_object('฿9,999.9'));
console.log(check_money_valid_old('฿9,999.9')); // true
console.log(set_money_object('฿9,999.9'));

console.log(check_money_valid_new('฿9,999')); // true
console.log(set_money_object('฿9,999'));
console.log(check_money_valid_old('฿9,999')); // true
console.log(set_money_object('฿9,999'));

//====
console.log("======"); // true

console.log(check_money_valid_new('9,999.99')); // true
console.log(set_money_object('9,999.99'));
console.log(check_money_valid_old('9,999.99')); // true
console.log(set_money_object('9,999.99'));

console.log(check_money_valid_new('9,999.999999999999')); // true
console.log(set_money_object('9,999.999999999999'));
console.log(check_money_valid_old('9,999.999999999999')); // false
console.log(set_money_object('9,999.999999999999'));

console.log(check_money_valid_new('9,999.9')); // true
console.log(set_money_object('9,999.9'));
console.log(check_money_valid_old('9,999.9')); // true
console.log(set_money_object('9,999.9'));

console.log(check_money_valid_new('9999.123456789')); // true
console.log(set_money_object('9999.123456789'));
console.log(check_money_valid_old('9999.123456789')); // true
console.log(set_money_object('9999.123456789'));

//====

