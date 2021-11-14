#pragma once

//typedef QStringList header;
class header : QStringList
{
	bool operator < (const header& theheader) const
	{
		if (this->first() != theheader.first())
		{
			return this->first() < theheader.first();
		}
	}
};

struct TABLE
{
	QString TableName;
	int     RowCount;
	int     ColCount;
	QList<header> Headers;
	QMap<header, QList<QVariant>> ContentData;

	QMap<QString, QStringList> _multilevelheader;
	// 可以考虑采用QmuitiMap

	TABLE()
	{
		TableName.clear();
		RowCount = 0;
		ColCount = 0;
		Headers.clear();
		ContentData.clear();
		_multilevelheader.clear();
	}

	void Clear()
	{
		TableName.clear();
		RowCount = 0;
		ColCount = 0;
		Headers.clear();
		ContentData.clear();
		_multilevelheader.clear();
	}
};