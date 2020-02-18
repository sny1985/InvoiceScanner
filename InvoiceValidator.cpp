#include "InvoiceValidator.h"
#include <QUrlQuery>
#include <QJsonDocument>
#include <QJsonObject>
#include <QJsonArray>

InvoiceValidator::InvoiceValidator()
{
    invoiceKeyModel.setColumnCount(6);
    invoiceKeyModel.setHeaderData(0, Qt::Horizontal, QString::fromLocal8Bit("发票代码"));
    invoiceKeyModel.setHeaderData(1, Qt::Horizontal, QString::fromLocal8Bit("发票号码"));
    invoiceKeyModel.setHeaderData(2, Qt::Horizontal, QString::fromLocal8Bit("开票日期"));
    invoiceKeyModel.setHeaderData(3, Qt::Horizontal, QString::fromLocal8Bit("合计金额"));
    invoiceKeyModel.setHeaderData(4, Qt::Horizontal, QString::fromLocal8Bit("校验码"));
    invoiceKeyModel.setHeaderData(5, Qt::Horizontal, QString::fromLocal8Bit("状态"));
}

InvoiceValidator::~InvoiceValidator()
{
    invoiceKeyModel.clear();
}

void InvoiceValidator::ReadPictureData(const QString &filePath)
{
    QFile file(filePath);
    if (file.open(QIODevice::ReadOnly))
    {
        char *buf = new char[file.size()];
        qint64 size = file.read(buf, file.size());
        if (file.size() == size)
        {
            pictureData.clear();
            EncodeBase64(buf, size, pictureData);
        }
        delete[] buf;

        file.close();
    }
}

void InvoiceValidator::AddTempItem(const QString &code, const QString &number, const QString &date, const QString &amount, const QString &check)
{
    int row = invoiceKeyModel.rowCount();
    invoiceKeyModel.setItem(row, 0, new QStandardItem(code));
    invoiceKeyModel.setItem(row, 1, new QStandardItem(number));
    invoiceKeyModel.setItem(row, 2, new QStandardItem(date));
    invoiceKeyModel.setItem(row, 3, new QStandardItem(amount));
    invoiceKeyModel.setItem(row, 4, new QStandardItem(check));
    invoiceKeyModel.setItem(row, 5, new QStandardItem(""));
}

void InvoiceValidator::OCR()
{
    // Post invoice picture
    pOCRRequest = new QNetworkRequest;
    pOCRNAManager = new QNetworkAccessManager(this);
    QMetaObject::Connection connRet = QObject::connect(pOCRNAManager, SIGNAL(finished(QNetworkReply *)), this, SLOT(onOCRRequestFinished(QNetworkReply *)));
    Q_ASSERT(connRet);

    QUrl url("https://netocr.com/api/recogInvoiveBase64.do");
    pOCRRequest->setHeader(QNetworkRequest::ContentTypeHeader, "application/x-www-form-urlencoded");
    pOCRRequest->setUrl(url);

    QUrlQuery query(url);
    query.addQueryItem("img", pictureData);
    query.addQueryItem("key", ocrKey);
    query.addQueryItem("secret", ocrSecret);
    query.addQueryItem("typeId", invoiceTypeId);
    query.addQueryItem("outvalue", "0");
    query.addQueryItem("format", "json");

    QByteArray data;
    data.append(query.query(QUrl::FullyEncoded));

    QNetworkReply *pReply = pOCRNAManager->post(*pOCRRequest, data);
}

void InvoiceValidator::ParseOCRResult(const QByteArray &jsonText)
{
    QJsonParseError jsonParserError;
    QJsonDocument jsonDoc = QJsonDocument::fromJson(jsonText, &jsonParserError);

    if (!jsonDoc.isNull() && jsonParserError.error == QJsonParseError::NoError)
    {
        if (jsonDoc.isObject())
        {
            InvoiceData invoiceTemp;

            QJsonObject jsonRootObj = jsonDoc.object();
            if (jsonRootObj.contains("cardsinfo"))
            {
                QJsonArray jsonCardArray = jsonRootObj.value("cardsinfo").toArray();
                for (auto i = 0; i < jsonCardArray.size(); ++i)
                {
                    if (jsonCardArray[i].isObject())
                    {
                        QJsonObject jsonCard = jsonCardArray[i].toObject();
                        if (jsonCard.contains("items"))
                        {
                            QJsonArray jsonItemArray = jsonCard.value("items").toArray();
                            for (auto j = 0; j < jsonItemArray.size(); ++j)
                            {
                                if (jsonItemArray[j].isObject())
                                {
                                    QJsonObject jsonItem = jsonItemArray[j].toObject();
                                    if (jsonItem.contains("desc"))
                                    {
                                        invoiceTemp.SetValue(jsonItem.value("desc").toString(), jsonItem.value("content").toString());
                                    }
                                }
                            }
                        }
                    }
                }
            }

            AddTempItem(invoiceTemp.GetInvoiceCode(), invoiceTemp.GetInvoiceNumber(), invoiceTemp.GetBillingDate(), invoiceTemp.GetTotalAmount(), invoiceTemp.GetCheckCode());
        }
    }
}

void InvoiceValidator::Verify()
{
    // Post invoice key data
    pVerifyRequest = new QNetworkRequest;
    pVerifyNAManager = new QNetworkAccessManager(this);
    QMetaObject::Connection connRet = QObject::connect(pVerifyNAManager, SIGNAL(finished(QNetworkReply *)), this, SLOT(onVerifyRequestFinished(QNetworkReply *)));
    Q_ASSERT(connRet);

    QUrl url("https://netocr.com/verapi/verInvoice.do");
    pVerifyRequest->setHeader(QNetworkRequest::ContentTypeHeader, "application/x-www-form-urlencoded");
    pVerifyRequest->setUrl(url);

    int row = invoiceKeyModel.rowCount() - 1;

    QUrlQuery query(url);
    query.addQueryItem("key", ocrKey);
    query.addQueryItem("secret", ocrSecret);
    query.addQueryItem("invoiceCode", invoiceKeyModel.item(row, 0)->text());
    query.addQueryItem("invoiceNumber", invoiceKeyModel.item(row, 1)->text());
    query.addQueryItem("billingDate", invoiceKeyModel.item(row, 2)->text());
    query.addQueryItem("totalAmount", invoiceKeyModel.item(row, 3)->text());
    query.addQueryItem("checkCode", invoiceKeyModel.item(row, 4)->text());
    query.addQueryItem("typeId", "3007");
    query.addQueryItem("format", "json");

    QByteArray data;
    data.append(query.query(QUrl::FullyEncoded));

    qDebug() << data;

    QNetworkReply *pReply = pVerifyNAManager->post(*pVerifyRequest, data);
}

void InvoiceValidator::ParseVerifyResult(const QByteArray &jsonText)
{
    QJsonParseError jsonParserError;
    QJsonDocument jsonDoc = QJsonDocument::fromJson(jsonText, &jsonParserError);

    if (!jsonDoc.isNull() && jsonParserError.error == QJsonParseError::NoError)
    {
        if (jsonDoc.isObject())
        {
            QJsonObject jsonRootObj = jsonDoc.object();
            if (jsonRootObj.contains("invoice"))
            {
                QJsonArray jsonInvoiceArray = jsonRootObj.value("invoice").toArray();
                for (auto i = 0; i < jsonInvoiceArray.size(); ++i)
                {
                    InvoiceData invoiceData;

                    if (jsonInvoiceArray[i].isObject())
                    {
                        QJsonObject jsonInvoice = jsonInvoiceArray[i].toObject();
                        if (jsonInvoice.contains("invoiceLists"))
                        {
                            QJsonArray jsonInvoiceListsArray = jsonInvoice.value("invoiceLists").toArray();
                            for (auto j = 0; j < jsonInvoiceListsArray.size(); ++j)
                            {
                                QList<InvoiceDetailData> invoiceList;

                                if (jsonInvoiceListsArray[j].isObject())
                                {
                                    QJsonObject jsonInvoiceLists = jsonInvoiceListsArray[j].toObject();
                                    if (jsonInvoiceLists.contains("invoiceList"))
                                    {
                                        QJsonArray jsonInvoiceListArray = jsonInvoiceLists.value("invoiceList").toArray();
                                        for (auto k = 0; k < jsonInvoiceListArray.size(); ++k)
                                        {
                                            if (jsonInvoiceListArray[k].isObject())
                                            {
                                                QJsonObject jsonInvoiceList = jsonInvoiceListArray[k].toObject();
                                                if (jsonInvoiceList.contains("veritem"))
                                                {
                                                    InvoiceDetailData detail;

                                                    QJsonArray jsonVerItemArray = jsonInvoiceList.value("veritem").toArray();
                                                    for (auto l = 0; l < jsonVerItemArray.size(); ++l)
                                                    {
                                                        if (jsonVerItemArray[l].isObject())
                                                        {
                                                            QJsonObject jsonVerItem = jsonVerItemArray[l].toObject();
                                                            if (jsonVerItem.contains("desc"))
                                                            {
                                                                detail.SetValue(jsonVerItem.value("desc").toString(), jsonVerItem.value("content").toString());
                                                            }
                                                        }
                                                    }
                                                    invoiceList.push_back(detail);
                                                }
                                            }
                                        }
                                    }
                                }

                                invoiceData.AddInvoiceList(invoiceList);
                            }
                        }
                        if (jsonInvoice.contains("veritem"))
                        {
                            QJsonArray jsonVerItemArray = jsonInvoice.value("veritem").toArray();
                            for (auto l = 0; l < jsonVerItemArray.size(); ++l)
                            {
                                if (jsonVerItemArray[l].isObject())
                                {
                                    QJsonObject jsonVerItem = jsonVerItemArray[l].toObject();
                                    if (jsonVerItem.contains("desc"))
                                    {
                                        invoiceData.SetValue(jsonVerItem.value("desc").toString(), jsonVerItem.value("content").toString());
                                    }
                                }
                            }
                        }
                    }

                    qDebug() << invoiceData;

                    mutex.lock();
                    invoiceDataList.push_back(invoiceData);
                    mutex.unlock();
                }
            }
        }
    }
}

//void InvoiceValidator::ImportFromExcel(const QString &filePath)
//{
//    QVariant excelData;
//    excel.Open(filePath);
//    excel.ReadCurrentWorkSheet(excelData);
//    excel.Close();

//    // Cast to QList<QList<QVariant>>
//    QList<QList<QVariant>> table;
//    QVariantList varRows = excelData.toList();
//    if (!varRows.isEmpty())
//    {
//        const int rowCount = varRows.size();
//        QVariantList rowData;
//        for (int i = 0; i < rowCount; ++i)
//        {
//            rowData = varRows[i].toList();
//            table.push_back(rowData);
//        }
//    }
//}

void InvoiceValidator::ExportToExcel(const QString &filePath, const QStringList &itemList)
{
    QList<QVariant> tableHead;
    InvoiceData::GetHead(tableHead, itemList);

    QList<QList<QVariant>> table;
    for (int i = 0; i < invoiceDataList.size(); ++i)
    {
        invoiceDataList.at(i).GetData(table, itemList);
    }

    QVariantList rows;
    int rowCount = table.size();
    rows.append(QVariant(tableHead));
    for (int i = 0; i < rowCount; ++i)
    {
        rows.append(QVariant(table[i]));
    }
    QVariant excelData = QVariant(rows);
    ++rowCount;

    QString colName;
    int i = tableHead.size();
    while (int mod = i % 26)
    {
        colName = char(mod + 0x40) + colName;
        i = i / 26;
    }
    QString excelDataRange = "A1:" + colName + QString::number(rowCount);

    QFileInfo fileInfo(filePath);
    if (!fileInfo.exists() || !fileInfo.isFile())
    {
        excel.Create(filePath);
    }
    else
    {
        excel.Open(filePath);
    }
    excel.WriteCurrentWorkSheet(excelData, excelDataRange);

    excel.Save();
    excel.Close();
}

void InvoiceValidator::onOCRRequestFinished(QNetworkReply *pReply)
{
    qDebug() << __FUNCTION__ << __FILE__ << __LINE__;

    QVariant statusCode = pReply->attribute(QNetworkRequest::HttpStatusCodeAttribute);
    if (statusCode.isValid())
    {
        qDebug() << "status code=" << statusCode.toInt();
    }

    QVariant reason = pReply->attribute(QNetworkRequest::HttpReasonPhraseAttribute).toString();
    if (reason.isValid())
    {
        qDebug() << "reason=" << reason.toString();
    }

    QNetworkReply::NetworkError err = pReply->error();
    if (err != QNetworkReply::NoError)
    {
        qDebug() << "Failed: " << pReply->errorString();
    }
    else
    {
        ParseOCRResult(pReply->readAll());

        Verify();
    }

    QObject::disconnect(pOCRNAManager);
    delete pOCRRequest;
    delete pOCRNAManager;
}

void InvoiceValidator::onVerifyRequestFinished(QNetworkReply *pReply)
{
    qDebug() << __FUNCTION__ << __FILE__ << __LINE__;

    QVariant statusCode = pReply->attribute(QNetworkRequest::HttpStatusCodeAttribute);
    if (statusCode.isValid())
    {
        qDebug() << "status code=" << statusCode.toInt();
    }

    QVariant reason = pReply->attribute(QNetworkRequest::HttpReasonPhraseAttribute).toString();
    if (reason.isValid())
    {
        qDebug() << "reason=" << reason.toString();
    }

    QNetworkReply::NetworkError err = pReply->error();
    if (err != QNetworkReply::NoError)
    {
        qDebug() << "Failed: " << pReply->errorString();
    }
    else
    {
        ParseVerifyResult(pReply->readAll());
    }

    QObject::disconnect(pVerifyNAManager);
    delete pVerifyRequest;
    delete pVerifyNAManager;
}

void InvoiceValidator::EncodeBase64(char *pData, qint64 len, QString &encoded)
{
    static const QString base64_chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

    unsigned char const *bytes_to_encode = reinterpret_cast<const unsigned char *>(pData);
    int i = 0;
    int j = 0;
    unsigned char char_array_3[3];
    unsigned char char_array_4[4];

    while (len--)
    {
        char_array_3[i++] = *(bytes_to_encode++);
        if (i == 3)
        {
            char_array_4[0] = (char_array_3[0] & 0xfc) >> 2;
            char_array_4[1] = ((char_array_3[0] & 0x03) << 4) + ((char_array_3[1] & 0xf0) >> 4);
            char_array_4[2] = ((char_array_3[1] & 0x0f) << 2) + ((char_array_3[2] & 0xc0) >> 6);
            char_array_4[3] = char_array_3[2] & 0x3f;

            for (i = 0; i < 4; ++i)
                encoded += base64_chars[char_array_4[i]];
            i = 0;
        }
    }

    if (i)
    {
        for (j = i; j < 3; ++j)
            char_array_3[j] = '/0';

        char_array_4[0] = (char_array_3[0] & 0xfc) >> 2;
        char_array_4[1] = ((char_array_3[0] & 0x03) << 4) + ((char_array_3[1] & 0xf0) >> 4);
        char_array_4[2] = ((char_array_3[1] & 0x0f) << 2) + ((char_array_3[2] & 0xc0) >> 6);
        char_array_4[3] = char_array_3[2] & 0x3f;

        for (j = 0; (j < i + 1); ++j)
            encoded += base64_chars[char_array_4[j]];

        while ((i++ < 3))
            encoded += '=';
    }

    encoded.replace(QRegExp("\\+"), "%2B");
}

QStandardItemModel * InvoiceValidator::GetModel()
{
    return &invoiceKeyModel;
}
