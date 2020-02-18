#ifndef INVOICEVALIDATOR_H
#define INVOICEVALIDATOR_H

#include <QNetworkRequest>
#include <QNetworkAccessManager>
#include <QNetworkReply>
#include <QMutex>
#include <QStandardItemModel>
#include "Invoice.h"
#include "ExcelHelper.h"

class InvoiceValidator : public QObject
{
    Q_OBJECT

public:
    InvoiceValidator();
    ~InvoiceValidator();

    void ReadPictureData(const QString &filePath);
    void AddTempItem(const QString &code, const QString &number, const QString &date, const QString &amount, const QString &check);
    void OCR();
    void ParseOCRResult(const QByteArray &jsonText);
    void Verify();
    void ParseVerifyResult(const QByteArray &jsonText);
//    void ImportFromExcel(const QString &filePath);
    void ExportToExcel(const QString &filePath, const QStringList &itemList);
    QStandardItemModel * GetModel();

private slots:
    void onOCRRequestFinished(QNetworkReply *pReply);
    void onVerifyRequestFinished(QNetworkReply *pReply);

private:
    QString ocrKey = "";
    QString ocrSecret = "";

    QNetworkRequest *pOCRRequest = nullptr;
    QNetworkAccessManager *pOCRNAManager = nullptr;
    QNetworkRequest *pVerifyRequest = nullptr;
    QNetworkAccessManager *pVerifyNAManager = nullptr;

    QMutex mutex;

    QString pictureData;
    QString invoiceTypeId = "2009";
    QList<InvoiceData> invoiceDataList;

    QStandardItemModel invoiceKeyModel;

    ExcelHelper excel;

    void EncodeBase64(char *pData, qint64 len, QString &encoded);
};

#endif // INVOICEVALIDATOR_H
