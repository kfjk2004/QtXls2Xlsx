// ����QBillProcess�����
#include <QtXlsx/QtXlsx>

int main()
{
    QXlsx::Document xlsx;
    xlsx.write("A1", "Hello Qt!");
    xlsx.saveAs("Test.xlsx");
    return 0;
}