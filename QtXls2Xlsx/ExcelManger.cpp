#include "ExcelManger.h"
#include <QDebug>
#if defined(Q_OS_WIN)
#include <QAxObject>
#include <windows.h>
#include <qdir.h>
#endif // Q_OS_WIN
#include <QVariant>


ExcelManger::ExcelManger(QObject* parent) : QObject(parent)
{
    // 在后台线程中使用QAxObject必须先初始化
    CoInitializeEx(NULL, COINIT_MULTITHREADED);
}

bool ExcelManger::Test(QString& path)
{
    QAxObject* excel = NULL;    //本例中，excel设定为Excel文件的操作对象
    QAxObject* workbooks = NULL;
    QAxObject* workbook = NULL;  //Excel操作对象
    excel = new QAxObject("Excel.Application");
    excel->dynamicCall("SetVisible(bool)", false); //true 表示操作文件时可见，false表示为不可见
    workbooks = excel->querySubObject("WorkBooks");


    //————————————————按文件路径打开文件————————————————————
    workbook = workbooks->querySubObject("Open(QString&)", path);
    // 获取打开的excel文件中所有的工作sheet
    QAxObject* worksheets = workbook->querySubObject("WorkSheets");


    //—————————————————Excel文件中表的个数:——————————————————
    int iWorkSheet = worksheets->property("Count").toInt();
    qDebug() << QString("Excel文件中表的个数: %1").arg(QString::number(iWorkSheet));


    // ————————————————获取第n个工作表 querySubObject("Item(int)", n);——————————
    QAxObject* worksheet = worksheets->querySubObject("Item(int)", 1);//本例获取第一个，最后参数填1


    //—————————获取该sheet的数据范围（可以理解为有数据的矩形区域）————
    QAxObject* usedrange = worksheet->querySubObject("UsedRange");

    //———————————————————获取行数———————————————
    QAxObject* rows = usedrange->querySubObject("Rows");
    int iRows = rows->property("Count").toInt();
    qDebug() << QString("行数为: %1").arg(QString::number(iRows));

    //————————————获取列数—————————
    QAxObject* columns = usedrange->querySubObject("Columns");
    int iColumns = columns->property("Count").toInt();
    qDebug() << QString("列数为: %1").arg(QString::number(iColumns));

    //————————数据的起始行———
    int iStartRow = rows->property("Row").toInt();
    qDebug() << QString("起始行为: %1").arg(QString::number(iStartRow));

    //————————数据的起始列————————————
    int iColumn = columns->property("Column").toInt();
    qDebug() << QString("起始列为: %1").arg(QString::number(iColumn));


    //——————————————读出数据—————————————
    //获取第i行第j列的数据
    //假如是第6行，第6列 对应表中F列6行，即F6
    QAxObject* range1 = worksheet->querySubObject("Range(QString)", "F6:F6");
    QString strRow6Col6 = "";
    strRow6Col6 = range1->property("Value").toString();
    qDebug() << "第6行，第6列的数据为：" + strRow6Col6;

    //待添加一个转换函数，即第6行，第6列，66转为F6


    //—————————————写入数据—————————————
    //获取F6的位置
    QAxObject* range2 = worksheet->querySubObject("Range(QString)", "F6:F6");
    //写入数据, 第6行，第6列
    range2->setProperty("Value", "中共十九大");
    QString newStr = "";
    newStr = range2->property("Value").toString();
    qDebug() << "写入数据后，第6行，第6列的数据为：" + newStr;

    //!!!!!!!一定要记得close，不然系统进程里会出现n个EXCEL.EXE进程
    //workbook->dynamicCall("Save()");
    QString temp = QDir::currentPath();
    temp += QString::fromLocal8Bit("/c.xlsx");
    // 有以下几种：
    //    51 = xlOpenXMLWorkbook(without macro’s in 2007 - 2016, xlsx) 保存为xlsx
    //    52 = xlOpenXMLWorkbookMacroEnabled(with or without macro’s in 2007 - 2016, xlsm) 保存为xlsm带宏的格式
    //    50 = xlExcel12(Excel Binary Workbook in 2007 - 2016 with or without macro’s, xlsb) 以二进制保存的工作表
    //    56 = xlExcel8(97 - 2003 format in Excel 2007 - 2016, xls) 以xls，2003格式保存的工作表
    //    ————————————————
    //    版权声明：本文为CSDN博主「viennadating」的原创文章，遵循CC 4.0 BY - SA版权协议，转载请附上原文出处链接及本声明。
    //    原文链接：https ://blog.csdn.net/viennadating/article/details/118066671
    workbook->dynamicCall("SaveAs(const QString&,int)", QDir::toNativeSeparators(temp),51);
    workbook->dynamicCall("Close()");
    excel->dynamicCall("Quit()");
    if (excel)
    {
        delete excel;
        excel = NULL;
    }

    return true;
}


ExcelManger::~ExcelManger()
{
}
