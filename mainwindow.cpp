#include "mainwindow.h"
#include "ui_mainwindow.h"

#include <QMessageBox>
#include <QFileDialog>
#include <QRegExp>
#include <QDate>

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);

    ui->tv_Invoices->setModel(iv.GetModel());
    ui->tv_Invoices->horizontalHeader()->setDefaultAlignment(Qt::AlignLeft);
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::on_btn_Open_clicked()
{
//    iv.Verify();

    QFileDialog openFileDialog(this);
    openFileDialog.setFileMode(QFileDialog::ExistingFiles);
    openFileDialog.setNameFilter(tr("Images (*.png *.bmp *.jpg)"));
    openFileDialog.setViewMode(QFileDialog::Detail);
    QStringList filePaths;
    if (openFileDialog.exec())
    {
        filePaths = openFileDialog.selectedFiles();
        directory = filePaths[0].left(filePaths[0].lastIndexOf(QRegExp("[\\/]")));
        for (int i = 0; i < filePaths.size(); ++i)
        {
            iv.ReadPictureData(filePaths.at(i));
            iv.OCR();
        }
    }
}

void MainWindow::on_btn_Add_clicked()
{
    iv.AddTempItem(ui->le_InvoiceCode->text(),
                   ui->le_InvoiceNumber->text(),
                   ui->le_BillingDate->text(),
                   ui->le_TotalAmount->text(),
                   ui->le_CheckCode->text());

    ui->tv_Invoices->setModel(iv.GetModel());

    iv.Verify();
}

void MainWindow::on_btn_Export_clicked()
{
//    QFile file1("data2.json");
//    if(!file1.open(QIODevice::ReadOnly))
//    {
//        QMessageBox::information(nullptr, "error", file1.errorString());
//    }

//    iv.ParseVerifyResult(file1.readAll());
//    file1.close();

    QDateTime datetime = QDateTime::currentDateTime();
    iv.ExportToExcel(QDir(directory).absoluteFilePath(datetime.toString("yyyyMMddhhmmss") + ".xlsx"),
                     QList<QString>({
                                        u8"发票号码",
                                        u8"开票日期",
                                        u8"购方名称",
                                        u8"销方名称",
                                        u8"货物或应税劳务名称",
                                        u8"规格型号",
                                        u8"数量",
                                        u8"单价",
                                        u8"税率",
                                    }));
}
