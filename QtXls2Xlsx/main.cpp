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
        qDebug() << QString::fromLocal8Bit("文件不存在");
        return a.exec();
    }
    else
    {
        //文件类型粗略检查
        if (!strPath.right(4).contains("xls"))
        {
            qDebug() << "只操作xlsx、xls文件";
            return a.exec();
        }
    }

    ExcelManger em;
    em.Test(strPath);

    FomatConvert convert;
    convert.read(strPath);
    convert.convert();

    qDebug()<< QString::fromLocal8Bit("执行完毕");

    return a.exec();
}
