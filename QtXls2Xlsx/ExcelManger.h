#pragma once

#include <QObject>

class ExcelManger : public QObject
{
	Q_OBJECT

public:

	explicit ExcelManger(QObject* parent = 0);

	bool Test(QString& path);

	~ExcelManger();
};
