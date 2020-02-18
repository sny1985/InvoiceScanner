#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include "InvoiceValidator.h"

QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

private slots:
    void on_btn_Open_clicked();
    void on_btn_Add_clicked();
    void on_btn_Export_clicked();

private:
    Ui::MainWindow *ui;

    InvoiceValidator iv;
    QString directory;
};
#endif // MAINWINDOW_H
