
// input str valide e164
function regexTestE164(str) {
    var regex = /^(\+)?[1-9]\d{1,14}$/;
    return regex.test(str);
}

//test
console.log(regexTestE164("+1234567890")); //true
console.log(regexTestE164("1234567890")); //true
console.log(regexTestE164("01234567890")); //false
console.log(regexTestE164("+01234567890")); //false
console.log(regexTestE164("1234567890+")); //false
console.log(regexTestE164("1234567890+1")); //false

