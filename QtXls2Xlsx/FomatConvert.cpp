#include "FomatConvert.h"
#include <QAxObject>
#include <windows.h>
#include <qdir.h>

FomatConvert::FomatConvert(QObject *parent)
	: QObject(parent)
{
	// 在后台线程中使用QAxObject必须先初始化
	CoInitializeEx(NULL, COINIT_MULTITHREADED);

	excel = NULL;			//本例中，excel设定为Excel文件的操作对象
	workbooks = NULL;
	workbook = NULL;		//Excel操作对象
	readfiletype = 0;
}

bool FomatConvert::read(const QString& filepath)
{
	excel = new QAxObject("Excel.Application");
	excel->dynamicCall("SetVisible(bool)", false); //true 表示操作文件时可见，false表示为不可见
	workbooks = excel->querySubObject("WorkBooks");


	//————————————————按文件路径打开文件————————————————————
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

    //!!!!!!!一定要记得close，不然系统进程里会出现n个EXCEL.EXE进程
    //workbook->dynamicCall("Save()");
    QString temp = QDir::currentPath();
    temp += c_filename;
    // 有以下几种：
    //    51 = xlOpenXMLWorkbook(without macro’s in 2007 - 2016, xlsx) 保存为xlsx
    //    52 = xlOpenXMLWorkbookMacroEnabled(with or without macro’s in 2007 - 2016, xlsm) 保存为xlsm带宏的格式
    //    50 = xlExcel12(Excel Binary Workbook in 2007 - 2016 with or without macro’s, xlsb) 以二进制保存的工作表
    //    56 = xlExcel8(97 - 2003 format in Excel 2007 - 2016, xls) 以xls，2003格式保存的工作表
    //    ————————————————
    //    版权声明：本文为CSDN博主「viennadating」的原创文章，遵循CC 4.0 BY - SA版权协议，转载请附上原文出处链接及本声明。
    //    原文链接：https ://blog.csdn.net/viennadating/article/details/118066671
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
