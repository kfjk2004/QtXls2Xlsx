#include <QCoreApplication>
#include <QFile>
#include <QDebug>
#include <QDir>
#include <qstring.h>
#include "ExcelManger.h"
#include "FomatConvert.h"

int main(int argc, char* argv[])
{
    QCoreApplication a(argc, argv);
    QString strPath = QDir::currentPath() + "/test.xlsx";
    QFile file(strPath);
    if (!file.exists())
    {
        qDebug() << QString::fromLocal8Bit("�ļ�������");
        return a.exec();
    }
    else
    {
        //�ļ����ʹ��Լ��
        if (!strPath.right(4).contains("xls"))
        {
            qDebug() << "ֻ����xlsx��xls�ļ�";
            return a.exec();
        }
    }

    ExcelManger em;
    em.Test(strPath);

    FomatConvert convert;
    convert.read(strPath);
    convert.convert();

    qDebug()<< QString::fromLocal8Bit("ִ�����");

    return a.exec();
}
