import std.array;
import std.algorithm;
import std.conv;
import std.datetime;
import std.exception;
import std.range;
import std.stdio;
import std.string;
import ae.sys.clipboard;

// 00 - Номер заявки 
// 01 - Код бумаги 
// 02 - Направление 
// 03 - МинДата и время заключения сделки 
// 04 - Кол-во ЦБ 
// 05 - Кол-во Номер сделки 
// 06 - Сумма сделки руб. 
// 07 - Сколько нашли
// 08 - Бот АВТО
// 09 - Стратегия АВТО
// 10 - Бот Залепа
// 11 - Стратегия Залепа
// 12 - Бот Итог
// 13 - Стратегия Итог
//13464857491	SBER	Купля	14.05.2015 15:05	1	1	747,8	0			213	TimeEnter	213	TimeEnter

void main() {
	int counter;
	string result;

	MyOrders[MyKey] orders;
	foreach(e; getClipboardText().splitter("\r\n").filter!"strip(a).length") {
		string[] cols = e.splitter("\t").array();

		enforce(cols.length == 14,
			format("Line with 14 cols req, have %s: %s", cols.length, cols));

		const keyo = MyKey(
			cols[1],
			cols[2],
			cols[4],
			cols[12]);
		const keyc = MyKey(
			cols[1],
			antiOper(cols[2]),
			cols[4],
			cols[12]);

		if (keyc in orders) { // Если есть что закрывать
			auto ordrs = orders[keyc];
			auto order = ordrs[0];
			ordrs = ordrs[1..$];
			if (ordrs.length)
				orders[keyc] = ordrs;
			else
				orders.remove(keyc);
			result ~= format("%s\t%s\t%s\r\n", 
				order.open, 
				order.n, 
				cleanUpOrder(e));
		} else {
			const tmp = MyOrder(cleanUpOrder(e), ++counter);
			if (keyo in orders)
				orders[keyo] ~= tmp;
			else
				orders[keyo] = [tmp];
		}
	}
	foreach(v; orders.byValue())
		foreach(e; v)
			result ~= format("%s\t%s\r\n", e.open, e.n);

	result = result
		.dup()
		.strip()
		.split("\r\n")
		.sort!(rowByIdCloseDate)
		.join("\r\n")
		.to!string();
	setClipboardText(result);
	writeln("Done OK");
}

bool rowByIdCloseDate(char[] a, char[] b) {
	auto aa = a.split("\t");
	auto bb = b.split("\t");
	int ia = (aa.length > 7) ? aa[7].to!int : 0;
	int ib = (bb.length > 7) ? bb[7].to!int : 0;
	string da = (aa.length > 13) ? aa[13].to!string : "01.01.5000 00:00";
	string db = (bb.length > 13) ? bb[13].to!string : "01.01.5000 00:00";
	if (ia == ib) {
		return excelStrToTime(da) < excelStrToTime(db);
	} else {
		return ia < ib;
	}
}

private:

enum sBuy = "Купля";
enum sSell = "Продажа";

struct MyKey {
	string instrument;
	string oper;
	string lot;
	string botName;
}

struct MyOrder {
	string open;
	int n;
}
alias MyOrders = MyOrder[];

string antiOper(in string oper) {
	switch (oper) {
		case sBuy:
			return sSell;
		case sSell:
			return sBuy;
		default:
			throw new Exception("Unknown oper " ~ oper);
	}
}

SysTime excelStrToTime(in string s) {
	// 29.05.15 12:00
	auto cols = s.replace(" ", ".")
		.replace(":", ".")
		.splitter(".")
		.array();
	if (cols.length < 5)
		throw new Exception("Invalid date-time '" ~ s ~ "'");
	int tmp = cols[2].to!int;
	return SysTime(DateTime(
		(tmp > 100)?tmp:tmp + 2000,
		cols[1].to!int,
		cols[0].to!int,
		cols[3].to!int,
		cols[4].to!int,
		0));
}

string cleanUpOrder(in string s) 
{
	auto tmp = s.splitter("\t");
	return chain(tmp.take(7), tmp.drop(12)).join("\t");
}
