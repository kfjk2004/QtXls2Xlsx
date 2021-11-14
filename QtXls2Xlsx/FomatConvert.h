#pragma once

#include <QObject>
#include <QAxObject>

class FomatConvert : public QObject
{
	Q_OBJECT

public:
	explicit FomatConvert(QObject *parent = 0);

	bool read(const QString& filepath);

	bool convert(bool _toXlsx = true);


	~FomatConvert();

public:
	QAxObject* excel;			//本例中，excel设定为Excel文件的操作对象
	QAxObject* workbooks;
	QAxObject* workbook;		//Excel操作对象

	int readfiletype;			// 0=初始化，1=xls格式，2=xlsx格式，3=csv格式
};
