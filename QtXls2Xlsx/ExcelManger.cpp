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
    // �ں�̨�߳���ʹ��QAxObject�����ȳ�ʼ��
    CoInitializeEx(NULL, COINIT_MULTITHREADED);
}

bool ExcelManger::Test(QString& path)
{
    QAxObject* excel = NULL;    //�����У�excel�趨ΪExcel�ļ��Ĳ�������
    QAxObject* workbooks = NULL;
    QAxObject* workbook = NULL;  //Excel��������
    excel = new QAxObject("Excel.Application");
    excel->dynamicCall("SetVisible(bool)", false); //true ��ʾ�����ļ�ʱ�ɼ���false��ʾΪ���ɼ�
    workbooks = excel->querySubObject("WorkBooks");


    //�����������������������������������ļ�·�����ļ�����������������������������������������
    workbook = workbooks->querySubObject("Open(QString&)", path);
    // ��ȡ�򿪵�excel�ļ������еĹ���sheet
    QAxObject* worksheets = workbook->querySubObject("WorkSheets");


    //����������������������������������Excel�ļ��б�ĸ���:������������������������������������
    int iWorkSheet = worksheets->property("Count").toInt();
    qDebug() << QString("Excel�ļ��б�ĸ���: %1").arg(QString::number(iWorkSheet));


    // ����������������������������������ȡ��n�������� querySubObject("Item(int)", n);��������������������
    QAxObject* worksheet = worksheets->querySubObject("Item(int)", 1);//������ȡ��һ������������1


    //��������������������ȡ��sheet�����ݷ�Χ���������Ϊ�����ݵľ������򣩡�������
    QAxObject* usedrange = worksheet->querySubObject("UsedRange");

    //����������������������������������������ȡ����������������������������������
    QAxObject* rows = usedrange->querySubObject("Rows");
    int iRows = rows->property("Count").toInt();
    qDebug() << QString("����Ϊ: %1").arg(QString::number(iRows));

    //��������������������������ȡ����������������������
    QAxObject* columns = usedrange->querySubObject("Columns");
    int iColumns = columns->property("Count").toInt();
    qDebug() << QString("����Ϊ: %1").arg(QString::number(iColumns));

    //�������������������ݵ���ʼ�С�����
    int iStartRow = rows->property("Row").toInt();
    qDebug() << QString("��ʼ��Ϊ: %1").arg(QString::number(iStartRow));

    //�������������������ݵ���ʼ�С�����������������������
    int iColumn = columns->property("Column").toInt();
    qDebug() << QString("��ʼ��Ϊ: %1").arg(QString::number(iColumn));


    //�����������������������������������ݡ�������������������������
    //��ȡ��i�е�j�е�����
    //�����ǵ�6�У���6�� ��Ӧ����F��6�У���F6
    QAxObject* range1 = worksheet->querySubObject("Range(QString)", "F6:F6");
    QString strRow6Col6 = "";
    strRow6Col6 = range1->property("Value").toString();
    qDebug() << "��6�У���6�е�����Ϊ��" + strRow6Col6;

    //�����һ��ת������������6�У���6�У�66תΪF6


    //��������������������������д�����ݡ�������������������������
    //��ȡF6��λ��
    QAxObject* range2 = worksheet->querySubObject("Range(QString)", "F6:F6");
    //д������, ��6�У���6��
    range2->setProperty("Value", "�й�ʮ�Ŵ�");
    QString newStr = "";
    newStr = range2->property("Value").toString();
    qDebug() << "д�����ݺ󣬵�6�У���6�е�����Ϊ��" + newStr;

    //!!!!!!!һ��Ҫ�ǵ�close����Ȼϵͳ����������n��EXCEL.EXE����
    //workbook->dynamicCall("Save()");
    QString temp = QDir::currentPath();
    temp += QString::fromLocal8Bit("/c.xlsx");
    // �����¼��֣�
    //    51 = xlOpenXMLWorkbook(without macro��s in 2007 - 2016, xlsx) ����Ϊxlsx
    //    52 = xlOpenXMLWorkbookMacroEnabled(with or without macro��s in 2007 - 2016, xlsm) ����Ϊxlsm����ĸ�ʽ
    //    50 = xlExcel12(Excel Binary Workbook in 2007 - 2016 with or without macro��s, xlsb) �Զ����Ʊ���Ĺ�����
    //    56 = xlExcel8(97 - 2003 format in Excel 2007 - 2016, xls) ��xls��2003��ʽ����Ĺ�����
    //    ��������������������������������
    //    ��Ȩ����������ΪCSDN������viennadating����ԭ�����£���ѭCC 4.0 BY - SA��ȨЭ�飬ת���븽��ԭ�ĳ������Ӽ���������
    //    ԭ�����ӣ�https ://blog.csdn.net/viennadating/article/details/118066671
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
