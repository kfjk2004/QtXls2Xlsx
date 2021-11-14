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
	QAxObject* excel;			//�����У�excel�趨ΪExcel�ļ��Ĳ�������
	QAxObject* workbooks;
	QAxObject* workbook;		//Excel��������

	int readfiletype;			// 0=��ʼ����1=xls��ʽ��2=xlsx��ʽ��3=csv��ʽ
};
