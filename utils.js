function moneyToString(_number, toUpper) {
	var toUpper = toUpper || false;
	var _arr_numbers = new Array();
	_arr_numbers[1] = new Array('', 'один', 'два', 'три', 'чотири', "п'ять", 'шість', 'сім', 'вісім', "дев'ять", 'десять', 'одинадцять', 'дванадцять', 'тринадцять ', ' чотирнадцять ', "п'ятнадцять", 'шістнадцять', 'сімнадцять', 'вісімнадцять', "дев'ятнадцять");
	_arr_numbers[2] = new Array('', '', 'двадцять', 'тридцять', 'сорок', "п'ятдесят", 'шістдесят', 'сімдесят', 'вісімдесят', "дев'яносто");
	_arr_numbers[3] = new Array('', 'сто', 'двісті', 'триста', 'чотириста', "п'ятсот", 'шістсот', 'сімсот', 'вісімсот', "дев'ятсот");
	function number_parser(_num, _desc) {
		var _string = '';
		var _num_hundred = '';
		if (_num.length == 3) {
			_num_hundred = _num.substr(0, 1);
			_num = _num.substr(1, 3);
			_string = _arr_numbers[3][_num_hundred] + ' ';
		}
		if (_num < 20) _string += _arr_numbers[1][parseFloat(_num)] + ' ';
		else {
			var _first_num = _num.substr(0, 1);
			var _second_num = _num.substr(1, 2);
			_string += _arr_numbers[2][_first_num] + ' ' + _arr_numbers[1][_second_num] + ' ';
		}
		switch (_desc) {
			case 0:
				if (_num.length == 2 && parseFloat(_num.substr(0, 1)) == 1) {
					_string += 'гривень';
					break;
				}
				var _last_num = parseFloat(_num.substr(-1));
				if (_last_num == 1) _string += 'гривня';
				else if (_last_num > 1 && _last_num < 5) _string += 'гривні';
				else _string += 'гривень';
				break;
			case 1:
				_num = _num.replace(/^[0]{1,}$/g, '0');
				if (_num.length == 2 && parseFloat(_num.substr(0, 1)) == 1) {
					_string += 'тисяч ';
					break;
				}
				var _last_num = parseFloat(_num.substr(-1));
				if (_last_num == 1) _string += 'тисяча ';
				else if (_last_num > 1 && _last_num < 5) _string += 'тисячі ';
				else if (parseFloat(_num) > 0) _string += 'тисяч ';
				_string = _string.replace('один ', 'одна ');
				_string = _string.replace('два ', 'дві ');
				break;
			case 2:
				_num = _num.replace(/^[0]{1,}$/g, '0');
				if (_num.length == 2 && parseFloat(_num.substr(0, 1)) == 1) {
					_string += 'мільйонів ';
					break;
				}
				var _last_num = parseFloat(_num.substr(-1));
				if (_last_num == 1) _string += 'мільйон ';
				else if (_last_num > 1 && _last_num < 5) _string += 'мільйони ';
				else if (parseFloat(_num) > 0) _string += 'мільйонів ';
				break;
			case 3:
				_num = _num.replace(/^[0]{1,}$/g, '0');
				if (_num.length == 2 && parseFloat(_num.substr(0, 1)) == 1) {
					_string += 'мільярдів ';
					break;
				}
				var _last_num = parseFloat(_num.substr(-1));
				if (_last_num == 1) _string += 'мільярд';
				else if (_last_num > 1 && _last_num < 5) _string += 'мільярда ';
				else if (parseFloat(_num) > 0) _string += 'мільярдів ';
				break;
		}
		return _string;
	}
	function decimals_parser(_num) {
		var _first_num = _num.substr(0, 1);
		var _second_num = parseFloat(_num.substr(1, 2));
		var _string = ' ' + _first_num + _second_num;
		if (_second_num == 1) _string += ' копійка';
		else if (_second_num > 1 && _second_num < 5) _string += ' копійки';
		else _string += ' копійок';
		return _string;
	}
	if (!_number || _number == 0) return false;
	if (typeof _number !== 'number') {
		_number = _number + '';
		_number = _number.replace(',', '.');
		_number = parseFloat(_number);
		if (isNaN(_number)) return false;
	}
	_number = _number.toFixed(2);
	if (_number.indexOf('.') != -1) {
		var _number_arr = _number.split('.');
		var _number = _number_arr[0];
		var _number_decimals = _number_arr[1];
	}
	var _number_length = _number.length;
	var _string = '';
	var _num_parser = '';
	var _count = 0;
	for (var _p = (_number_length - 1); _p >= 0; _p--) {
		var _num_digit = _number.substr(_p, 1);
		_num_parser = _num_digit + _num_parser;
		if ((_num_parser.length == 3 || _p == 0) && !isNaN(parseFloat(_num_parser))) {
			_string = number_parser(_num_parser, _count) + _string;
			_num_parser = '';
			_count++;
		}
	}
	if (_number_decimals) _string += decimals_parser(_number_decimals);
	if (toUpper === true || toUpper == 'upper') {
		_string = _string.substr(0, 1).toUpperCase() + _string.substr(1);
	}
	return _string.replace(/[\s]{1,}/g, ' ');
};

function numberToString(_number, isFixed, toUpper) {
	var toUpper = toUpper || false;
	var _arr_numbers = new Array();
	_arr_numbers[1] = new Array('', 'один', 'два', 'три', 'чотири', "п'ять", 'шість', 'сім', 'вісім', "дев'ять", 'десять', 'одинадцять', 'дванадцять', 'тринадцять ', ' чотирнадцять ', "п'ятнадцять", 'шістнадцять', 'сімнадцять', 'вісімнадцять', "дев'ятнадцять");
	_arr_numbers[2] = new Array('', '', 'двадцять', 'тридцять', 'сорок', "п'ятдесят", 'шістдесят', 'сімдесят', 'вісімдесят', "дев'яносто");
	_arr_numbers[3] = new Array('', 'сто', 'двісті', 'триста', 'чотириста', "п'ятсот", 'шістсот', 'сімсот', 'вісімсот', "дев'ятсот");

	function number_parser(_num, _desc) {
		var _string = '';
		var _num_hundred = '';
		if (_num.length == 3) {
			_num_hundred = _num.substr(0, 1);
			_num = _num.substr(1, 3);
			_string = _arr_numbers[3][_num_hundred] + ' ';
		}
		if (_num < 20) _string += _arr_numbers[1][parseFloat(_num)] + ' ';
		else {
			var _first_num = _num.substr(0, 1);
			var _second_num = _num.substr(1, 2);
			_string += _arr_numbers[2][_first_num] + ' ' + _arr_numbers[1][_second_num] + ' ';
		}
		switch (_desc) {
			case 0:
				if (_num.length == 2 && parseFloat(_num.substr(0, 1)) == 1) {
					_string += '';
					break;
				}
				var _last_num = parseFloat(_num.substr(-1));
				if (_last_num == 1) _string += '';
				else if (_last_num > 1 && _last_num < 5) _string += '';
				else _string += '';
				break;
			case 1:
				_num = _num.replace(/^[0]{1,}$/g, '0');
				if (_num.length == 2 && parseFloat(_num.substr(0, 1)) == 1) {
					_string += 'тисяч ';
					break;
				}
				var _last_num = parseFloat(_num.substr(-1));
				if (_last_num == 1) _string += 'тисяча ';
				else if (_last_num > 1 && _last_num < 5) _string += 'тисячі ';
				else if (parseFloat(_num) > 0) _string += 'тисяч ';
				_string = _string.replace('один ', 'одна ');
				_string = _string.replace('два ', 'дві ');
				break;
			case 2:
				_num = _num.replace(/^[0]{1,}$/g, '0');
				if (_num.length == 2 && parseFloat(_num.substr(0, 1)) == 1) {
					_string += 'мільйонів ';
					break;
				}
				var _last_num = parseFloat(_num.substr(-1));
				if (_last_num == 1) _string += 'мільйон ';
				else if (_last_num > 1 && _last_num < 5) _string += 'мільйони ';
				else if (parseFloat(_num) > 0) _string += 'мільйонів ';
				break;
			case 3:
				_num = _num.replace(/^[0]{1,}$/g, '0');
				if (_num.length == 2 && parseFloat(_num.substr(0, 1)) == 1) {
					_string += 'мільярдів ';
					break;
				}
				var _last_num = parseFloat(_num.substr(-1));
				if (_last_num == 1) _string += 'мільярд';
				else if (_last_num > 1 && _last_num < 5) _string += 'мільярда ';
				else if (parseFloat(_num) > 0) _string += 'мільярдів ';
				break;
		}
		return _string;
	}
	function decimals_parser(_num) {
		var _first_num = _num.substr(0, 1);
		var _second_num = parseFloat(_num.substr(1, 2));
		var _string = ' ' + _first_num + _second_num;
		if (_second_num == 1) _string += '';
		else if (_second_num > 1 && _second_num < 5) _string += '';
		else _string += '';
		return _string;
	}
	if (!_number || _number == 0) return false;
	if (typeof _number !== 'number') {
		_number = _number + '';
		_number = _number.replace(',', '.');
		_number = parseFloat(_number);
		if (isNaN(_number)) return false;
	}
	_number = _number.toFixed(2);
	if (_number.indexOf('.') != -1) {
		var _number_arr = _number.split('.');
		var _number = _number_arr[0];
		var _number_decimals = _number_arr[1];
	}
	var _number_length = _number.length;
	var _string = '';
	var _num_parser = '';
	var _count = 0;
	for (var _p = (_number_length - 1); _p >= 0; _p--) {
		var _num_digit = _number.substr(_p, 1);
		_num_parser = _num_digit + _num_parser;
		if ((_num_parser.length == 3 || _p == 0) && !isNaN(parseFloat(_num_parser))) {
			_string = number_parser(_num_parser, _count) + _string;
			_num_parser = '';
			_count++;
		}
	}
	if (_number_decimals && isFixed) _string += decimals_parser(_number_decimals);
	if (toUpper === true || toUpper == 'upper') {
		_string = _string.substr(0, 1).toUpperCase() + _string.substr(1);
	}
	return _string.replace(/[\s]{1,}/g, ' ');
};

function numberToDigit(number) {
	return String(number).replace(/(\d)(?=(\d\d\d)+([^\d]|$))/g, '$1 ');
}

function writeLog(text) {
	const fs = require("fs");
}

function formatDate(date) {

	var dd = date.getDate();
	if (dd < 10) dd = '0' + dd;

	var mm = date.getMonth() + 1;
	if (mm < 10) mm = '0' + mm;

	var yy = date.getFullYear() % 100;
	if (yy < 10) yy = '0' + yy;

	return dd + '.' + mm + '.' + '20' + yy;
}


Number.prototype.numberToString = function (toUpper) {
	return numberToString(this, toUpper);
};
String.prototype.numberToString = function (toUpper) {
	return numberToString(this, toUpper);
};

Number.prototype.moneyToString = function (toUpper) {
	return moneyToString(this, toUpper);
};
String.prototype.moneyToString = function (toUpper) {
	return moneyToString(this, toUpper);
};

module.exports = {
	numberToString,
	moneyToString,
	numberToDigit,
	formatDate,
	writeLog
}