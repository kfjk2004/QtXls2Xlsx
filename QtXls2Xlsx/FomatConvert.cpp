#include "FomatConvert.h"
#include <QAxObject>
#include <windows.h>
#include <qdir.h>

FomatConvert::FomatConvert(QObject *parent)
	: QObject(parent)
{
	// �ں�̨�߳���ʹ��QAxObject�����ȳ�ʼ��
	CoInitializeEx(NULL, COINIT_MULTITHREADED);

	excel = NULL;			//�����У�excel�趨ΪExcel�ļ��Ĳ�������
	workbooks = NULL;
	workbook = NULL;		//Excel��������
	readfiletype = 0;
}

bool FomatConvert::read(const QString& filepath)
{
	excel = new QAxObject("Excel.Application");
	excel->dynamicCall("SetVisible(bool)", false); //true ��ʾ�����ļ�ʱ�ɼ���false��ʾΪ���ɼ�
	workbooks = excel->querySubObject("WorkBooks");


	//�����������������������������������ļ�·�����ļ�����������������������������������������
	workbook = workbooks->querySubObject("Open(QString&)", filepath);

    if (NULL != workbook)
    {
        return true;
    }
	return false;
}

bool FomatConvert::convert(bool _toXlsx)
{
    QString c_filename;
    int c_fileformat;
    if (_toXlsx)
    {
        c_filename = QString::fromLocal8Bit("/convert.xlsx");
        c_fileformat = 51;
    }
    else
    {
        c_filename = QString::fromLocal8Bit("/convert.xls");
        c_fileformat = 56;
    }

    //!!!!!!!һ��Ҫ�ǵ�close����Ȼϵͳ����������n��EXCEL.EXE����
    //workbook->dynamicCall("Save()");
    QString temp = QDir::currentPath();
    temp += c_filename;
    // �����¼��֣�
    //    51 = xlOpenXMLWorkbook(without macro��s in 2007 - 2016, xlsx) ����Ϊxlsx
    //    52 = xlOpenXMLWorkbookMacroEnabled(with or without macro��s in 2007 - 2016, xlsm) ����Ϊxlsm����ĸ�ʽ
    //    50 = xlExcel12(Excel Binary Workbook in 2007 - 2016 with or without macro��s, xlsb) �Զ����Ʊ���Ĺ�����
    //    56 = xlExcel8(97 - 2003 format in Excel 2007 - 2016, xls) ��xls��2003��ʽ����Ĺ�����
    //    ��������������������������������
    //    ��Ȩ����������ΪCSDN������viennadating����ԭ�����£���ѭCC 4.0 BY - SA��ȨЭ�飬ת���븽��ԭ�ĳ������Ӽ���������
    //    ԭ�����ӣ�https ://blog.csdn.net/viennadating/article/details/118066671
    workbook->dynamicCall("SaveAs(const QString&,int)", QDir::toNativeSeparators(temp), c_fileformat);
    workbook->dynamicCall("Close()");
    excel->dynamicCall("Quit()");
    if (excel)
    {
        delete excel;
        excel = NULL;
    }
	return true;
}

FomatConvert::~FomatConvert()
{
}
