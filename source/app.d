import std.stdio;
import std.string;
import std.array;
import std.algorithm;
import std.conv;
import std.datetime;
import std.range;
import ae.sys.clipboard;

//00-Номер заявки
//01-Код бумаги
//02-Направление
//03-МинДата и время заключения сделки
//04-Кол-во ЦБ
//05-Кол-во
//06-Номер сделки 
//07-Сумма сделки расчет
//08-Количество ботов
//09-Имя бота
//10-Стратегия
//13464857491	SBER	Купля	14.05.15 15:05	1	1	747,8
void main() {
	int counter;
	string result;

	MyOrders[MyKey] orders;
	foreach(e; getClipboardText()[0..$-1].splitter("\r\n").filter!"strip(a).length") {
		string[] cols = e.splitter("\t").array();

		if (cols.length <= 4)
			throw new Exception("Line too short: " ~ e);

		const keyo = MyKey(cols[1],         (cols[2]), cols[4], (cols.length > 9)?cols[9]:"");
		const keyc = MyKey(cols[1], antiOper(cols[2]), cols[4], (cols.length > 9)?cols[9]:"");

		if (keyc in orders) { // Если есть что закрывать
			auto ordrs = orders[keyc];
			auto order = ordrs[0];
			ordrs = ordrs[1..$];
			if (ordrs.length)
				orders[keyc] = ordrs;
			else
				orders.remove(keyc);
			result ~= format("%s\t%s\topen\t%s\r\n", order.open, order.n, e);
		} else {
			const tmp = MyOrder(e, ++counter);
			if (keyo in orders)
				orders[keyo] ~= tmp;
			else
				orders[keyo] = [tmp];
		}
	}
	foreach(v; orders.byValue())
		foreach(e; v)
			result ~= format("%s\t%s\topened\r\n", e.open, e.n);

	setClipboardText(result);
	std.file.write("dest-orders.txt", result);
	writeln("Done");
	readln();
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
	if (cols.length <= 5)
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

