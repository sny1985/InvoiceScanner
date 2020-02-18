#ifndef INVOICE_H
#define INVOICE_H

#include <QString>
#include <QVector>
#include <QHash>
#include <QDebug>
#include "ExcelHelper.h"

struct DataProperty
{
    QString key;
    NumberFormat numberFormat;
};

class InvoiceDetailData
{
public:
    InvoiceDetailData();
    ~InvoiceDetailData();

    void SetValue(const QString &name, const QString &value);

    friend QDebug operator<<(QDebug d, const InvoiceDetailData &id);

    friend class InvoiceData;

private:
    QVector<QString> invoiceDetailItems;

    static QVector<DataProperty> invoiceDetailItemProperties;
    static QHash<QString, int> invoiceDetailItemIndices;
};

class InvoiceData
{
public:
    InvoiceData();
    ~InvoiceData();

    void SetValue(const QString &name, const QString &value);
    void AddInvoiceList(const QList<InvoiceDetailData> &invoiceList);
    const QString & GetInvoiceCode() const;
    const QString & GetInvoiceNumber() const;
    const QString & GetBillingDate() const;
    const QString & GetTotalAmount() const;
    const QString & GetCheckCode() const;
    void GetData(QList<QList<QVariant>> &output, const QStringList &itemList) const;

    static void GetHead(QList<QVariant> &output, const QStringList &itemList);

    friend QDebug operator<<(QDebug d, const InvoiceData &id);

private:
    QVector<QString> invoiceItems;
    QList<QList<InvoiceDetailData>> details;

    static QVector<DataProperty> invoiceItemProperties;
    static QHash<QString, int> invoiceItemIndices;
};

inline QDebug operator<<(QDebug d, const InvoiceDetailData &idd)
{
    for (int i = 0; i < InvoiceDetailData::invoiceDetailItemProperties.size(); ++i)
    {
        d << InvoiceDetailData::invoiceDetailItemProperties[i].key
          << ": " << idd.invoiceDetailItems[i] << endl;
    }

    return d;
}

inline QDebug operator<<(QDebug d, const InvoiceData &id)
{
    for (int i = 0; i < InvoiceData::invoiceItemProperties.size(); ++i)
    {
        d << InvoiceData::invoiceItemProperties[i].key
          << ": " << id.invoiceItems[i] << endl;
    }

    for (int i = 0; i < id.details.size(); ++i)
    {
        for (int j = 0; j < id.details.at(i).size(); ++j)
        {
            d << id.details.at(i).at(j);
        }
    }

    return d;
}

#endif // INVOICE_H
