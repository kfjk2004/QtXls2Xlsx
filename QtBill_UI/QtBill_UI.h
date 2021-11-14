#pragma once

#include <QtWidgets/QMainWindow>
#include "ui_QtBill_UI.h"

class QtBill_UI : public QMainWindow
{
    Q_OBJECT

public:
    QtBill_UI(QWidget *parent = Q_NULLPTR);

private:
    Ui::QtBill_UIClass ui;
};
