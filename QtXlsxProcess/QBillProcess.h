#pragma once

#include <QObject>
#include <QtXlsx>
#include <QList>
#include <QMap>

#include "Table_Interface.h"

// TODO 目前默认仅读取指定index的一个sheet表

class QBillProcess : public QObject
{
	Q_OBJECT

public:
	QBillProcess(QObject *parent);
	~QBillProcess();

public:
	TABLE m_Table;
	TABLE m_Lookup;

public:
	// 文件操作功能
	// 文件读取和保存功能
	bool f_read_file(QString absfilepath);
	bool f_save_to_file(QString absfilepath);
	// 查询功能
	// 设置查询条件和查询函数
	void f_set_lookupcondition(QStringList headers);
	void f_set_lookupcondition(QString headers,int look_value);
	void f_set_lookupcondition(QString headers, QString look_string);
	bool f_lookup();

	

private:
	// 内部数据
	QList<QStringList> _raw_data;

private:
	// 配置参数
	int				   _header_level;
	int				   _sheet_index;
	QString			   _convert_xlsx_filepath;
	QString            _convert_xlsx_filename;

private:
	// 自定义业务函数
	bool _f_fill_Table();
	bool _f_wash_data();

	bool _f_lookup(QList<QVariant> lookup_condition);
	// 查询条件：统一用QVariant表示的内容，包括：1-header 表头字符串、2-Qmap<header，Qvariant>-某表头代表的某一列等与Qvariant的值的查询

private:
	// 自定义处理函数
	bool _cf_drop_whitespace(QStringList& data);
	bool _cf_convert_xls2xlsx(QString absfilepath);
};
