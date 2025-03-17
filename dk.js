function setNumber(moneyFiled){
	var x = moneyFiled.replace("฿", "");
	var y = x.replace(",", "");
	return y;
}

setNumber("");
