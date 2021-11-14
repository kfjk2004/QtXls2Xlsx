#pragma once

#include <QObject>
#include <QtXlsx>
#include <QList>
#include <QMap>

#include "Table_Interface.h"

// TODO ĿǰĬ�Ͻ���ȡָ��index��һ��sheet��

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
	// �ļ���������
	// �ļ���ȡ�ͱ��湦��
	bool f_read_file(QString absfilepath);
	bool f_save_to_file(QString absfilepath);
	// ��ѯ����
	// ���ò�ѯ�����Ͳ�ѯ����
	void f_set_lookupcondition(QStringList headers);
	void f_set_lookupcondition(QString headers,int look_value);
	void f_set_lookupcondition(QString headers, QString look_string);
	bool f_lookup();

	

private:
	// �ڲ�����
	QList<QStringList> _raw_data;

private:
	// ���ò���
	int				   _header_level;
	int				   _sheet_index;
	QString			   _convert_xlsx_filepath;
	QString            _convert_xlsx_filename;

private:
	// �Զ���ҵ����
	bool _f_fill_Table();
	bool _f_wash_data();

	bool _f_lookup(QList<QVariant> lookup_condition);
	// ��ѯ������ͳһ��QVariant��ʾ�����ݣ�������1-header ��ͷ�ַ�����2-Qmap<header��Qvariant>-ĳ��ͷ�����ĳһ�е���Qvariant��ֵ�Ĳ�ѯ

private:
	// �Զ��崦����
	bool _cf_drop_whitespace(QStringList& data);
	bool _cf_convert_xls2xlsx(QString absfilepath);
};
