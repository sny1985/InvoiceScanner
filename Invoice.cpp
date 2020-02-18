#include "Invoice.h"

QVector<DataProperty> InvoiceDetailData::invoiceDetailItemProperties = {
    {u8"货物或应税劳务名称", NumberFormat::none},
    {u8"规格型号", NumberFormat::none},
    {u8"单位", NumberFormat::none},
    {u8"数量", NumberFormat::general},
    {u8"单价", NumberFormat::currency},
    {u8"金额", NumberFormat::currency},
    {u8"税率", NumberFormat::percentage},
    {u8"税额", NumberFormat::currency},
};

QHash<QString, int> InvoiceDetailData::invoiceDetailItemIndices = {
    {u8"货物或应税劳务名称", 0},
    {u8"规格型号", 1},
    {u8"单位", 2},
    {u8"数量", 3},
    {u8"单价", 4},
    {u8"金额", 5},
    {u8"税率", 6},
    {u8"税额", 7},
};

QVector<DataProperty> InvoiceData::invoiceItemProperties = {
    {u8"发票类型", NumberFormat::none},
    {u8"所属行政区名称", NumberFormat::none},
    {u8"发票代码", NumberFormat::text},
    {u8"发票号码", NumberFormat::text},
    {u8"开票日期", NumberFormat::none},
    {u8"购方名称", NumberFormat::none},
    {u8"购方税号", NumberFormat::text},
    {u8"购方地址电话", NumberFormat::text},
    {u8"购方开户行账号", NumberFormat::none},
    {u8"销方名称", NumberFormat::none},
    {u8"销方税号", NumberFormat::text},
    {u8"销方地址电话", NumberFormat::text},
    {u8"销方开户行账号", NumberFormat::text},
    {u8"合计金额", NumberFormat::currency},
    {u8"合计税额", NumberFormat::currency},
    {u8"价税合计", NumberFormat::currency},
    {u8"价税合计_中文", NumberFormat::none},
    {u8"备注", NumberFormat::none},
    {u8"机器编号", NumberFormat::text},
    {u8"校验码", NumberFormat::text},
    {u8"发票状态", NumberFormat::none},
    {u8"明细内容", NumberFormat::none},
    {u8"大写金额", NumberFormat::none},
    {u8"收款人", NumberFormat::none},
    {u8"复核", NumberFormat::none},
    {u8"开票人", NumberFormat::none},
    {u8"发票联", NumberFormat::none},
    {u8"印刷发票代码", NumberFormat::text},
    {u8"印刷发票号码", NumberFormat::text},
    {u8"二维码发票代码", NumberFormat::text},
    {u8"二维码发票号码", NumberFormat::text},
    {u8"发票综合代码", NumberFormat::text},
    {u8"发票综合号码", NumberFormat::text},
    {u8"是否有印章", NumberFormat::none},
    //        {u8"二维码", NumberFormat::text},
    //        {u8"密文", NumberFormat::none},
    //        {u8"票据属性", NumberFormat::none},
    //        {u8"指定区域二维码", NumberFormat::text},
    //        {u8"承运人名称", NumberFormat::none},
    //        {u8"承运人识别号", NumberFormat::text},
    //        {u8"受票方名称",  NumberFormat::none},
    //        {u8"受票方识别号", NumberFormat::text},
    //        {u8"运输货物信息", NumberFormat::none},
    //        {u8"起运地、经由、到达地", NumberFormat::none},
    //        {u8"税控盘号", NumberFormat::text},
    //        {u8"车种车号", NumberFormat::text},
    //        {u8"车船吨位", NumberFormat::none},
    //        {u8"主管税务机关", NumberFormat::none},
    //        {u8"主管税务名称", NumberFormat::none},
    //        {u8"身份证号码", NumberFormat::text},
    //        {u8"车辆类型", NumberFormat::none},
    //        {u8"厂牌型号", NumberFormat::text},
    //        {u8"产地", NumberFormat::none},
    //        {u8"合格证号", NumberFormat::text},
    //        {u8"商检单号", NumberFormat::text},
    //        {u8"发动机号", NumberFormat::text},
    //        {u8"车辆识别代号", NumberFormat::text},
    //        {u8"进口证明书号", NumberFormat::text},
    //        {u8"完税凭证号码", NumberFormat::text},
    //        {u8"限乘人数", NumberFormat::general},
};

QHash<QString, int> InvoiceData::invoiceItemIndices = {
    {u8"发票类型", 0},
    {u8"所属行政区名称", 1},
    {u8"发票代码", 2},
    {u8"发票号码", 3},
    {u8"开票日期", 4},
    {u8"购方名称", 5},
    {u8"购方税号", 6},
    {u8"购方地址、电话", 7},
    {u8"购方地址电话", 7},
    {u8"购方开户行及账号", 8},
    {u8"购方开户行账号", 8},
    {u8"销方名称", 9},
    {u8"销方税号", 10},
    {u8"销方地址、电话", 11},
    {u8"销方地址电话", 11},
    {u8"销方开户行及账号", 12},
    {u8"销方开户行账号", 12},
    {u8"合计金额", 13},
    {u8"合计税额", 14},
    {u8"价税合计", 15},
    {u8"价税合计_中文", 16},
    {u8"备注", 17},
    {u8"机器编号", 18},
    {u8"校验码", 19},
    {u8"发票状态", 20},
    {u8"明细内容", 21},
    {u8"大写金额", 22},
    {u8"收款人", 23},
    {u8"复核", 24},
    {u8"开票人", 25},
    {u8"发票联", 26},
    {u8"印刷发票代码", 27},
    {u8"印刷发票号码", 28},
    {u8"二维码发票代码", 29},
    {u8"二维码发票号码", 30},
    {u8"发票综合代码", 31},
    {u8"发票综合号码", 32},
    {u8"是否有印章", 33},
//    {u8"二维码", 34},
//    {u8"密文", 35},
//    {u8"票据属性", 36},
//    {u8"指定区域二维码", 37},
//    {u8"承运人名称", 38},
//    {u8"承运人识别号", 39},
//    {u8"受票方名称", 40},
//    {u8"受票方识别号", 41},
//    {u8"运输货物信息", 42},
//    {u8"起运地、经由、到达地", 43},
//    {u8"税控盘号", 44},
//    {u8"车种车号", 45},
//    {u8"车船吨位", 46},
//    {u8"主管税务机关", 47},
//    {u8"主管税务名称", 48},
//    {u8"身份证号码", 49},
//    {u8"车辆类型", 50},
//    {u8"厂牌型号", 51},
//    {u8"产地", 52},
//    {u8"合格证号", 53},
//    {u8"商检单号", 54},
//    {u8"发动机号", 55},
//    {u8"车辆识别代号", 56},
//    {u8"进口证明书号", 57},
//    {u8"完税凭证号码", 58},
//    {u8"限乘人数", 59},
};

InvoiceDetailData::InvoiceDetailData()
{
    invoiceDetailItems.resize(invoiceDetailItemProperties.size());
}

InvoiceDetailData::~InvoiceDetailData()
{
    invoiceDetailItems.clear();
}

void InvoiceDetailData::SetValue(const QString &name, const QString &value)
{
    if (invoiceDetailItemIndices.contains(name))
    {
        invoiceDetailItems[invoiceDetailItemIndices[name]] = value;
    }
}

InvoiceData::InvoiceData()
{
    invoiceItems.resize(invoiceItemProperties.size());
}

InvoiceData::~InvoiceData()
{
    invoiceItems.clear();

    for (int i = 0; i < details.size(); ++i)
    {
        details[i].clear();
    }
    details.clear();
}

void InvoiceData::SetValue(const QString &name, const QString &value)
{
    if (invoiceItemIndices.contains(name))
    {
        invoiceItems[invoiceItemIndices[name]] = value;
    }
}

void InvoiceData::AddInvoiceList(const QList<InvoiceDetailData> &invoiceList)
{
    details.push_back(invoiceList);
}

const QString & InvoiceData::GetInvoiceCode() const
{
    return invoiceItems[invoiceItemIndices[u8"发票代码"]];
}

const QString & InvoiceData::GetInvoiceNumber() const
{
    return  invoiceItems[invoiceItemIndices[u8"发票号码"]];
}

const QString & InvoiceData::GetBillingDate() const
{
    return invoiceItems[invoiceItemIndices[u8"开票日期"]];
}

const QString & InvoiceData::GetTotalAmount() const
{
    return invoiceItems[invoiceItemIndices[u8"合计金额"]];
}

const QString & InvoiceData::GetCheckCode() const
{
    return invoiceItems[invoiceItemIndices[u8"校验码"]];
}

void InvoiceData::GetData(QList<QList<QVariant>> &output, const QStringList &itemList) const
{
    QStringList itemNames = itemList;
    if (itemNames.empty())
    {
        for (int i = 0; i < invoiceItemProperties.size(); ++i)
        {
            if (InvoiceData::invoiceItemProperties[i].key != u8"明细内容")
            {
                itemNames.push_back(InvoiceData::invoiceItemProperties[i].key);
            }
            else
            {
                for (int l = 0; l < InvoiceDetailData::invoiceDetailItemProperties.size(); ++l)
                {
                    itemNames.push_back(InvoiceDetailData::invoiceDetailItemProperties[i].key);
                }
            }
        }
    }
    if (details.empty())
    {
        QList<QVariant> row;
        for (int i = 0; i < itemNames.size(); ++i)
        {
            if (InvoiceData::invoiceItemIndices.contains(itemNames[i]))
            {
                row.push_back(invoiceItems[InvoiceData::invoiceItemIndices.value(itemNames[i])]);
            }
            else
            {
                row.push_back("");
            }
        }
        output.push_back(row);
    }
    else
    {
        for (int i = 0; i < details.size(); ++i)
        {
            for (int j = 0; j < details.at(i).size(); ++j)
            {
                const InvoiceDetailData &detail = details.at(i).at(j);
                QList<QVariant> row;
                for (int k = 0; k < itemNames.size(); ++k)
                {
                    if (InvoiceData::invoiceItemIndices.contains(itemNames[k]))
                    {
                        row.push_back(invoiceItems[InvoiceData::invoiceItemIndices.value(itemNames[k])]);
                    }
                    if (InvoiceDetailData::invoiceDetailItemIndices.contains(itemNames[k]))
                    {
                        row.push_back(detail.invoiceDetailItems[InvoiceDetailData::invoiceDetailItemIndices.value(itemNames[k])]);
                    }
                }
                output.push_back(row);
            }
        }
    }
}

void InvoiceData::GetHead(QList<QVariant> &output, const QStringList &itemList)
{
    QStringList itemNames = itemList;
    if (itemNames.empty())
    {
        for (int i = 0; i < invoiceItemProperties.size(); ++i)
        {
            if (InvoiceData::invoiceItemProperties[i].key != u8"明细内容")
            {
                itemNames.push_back(InvoiceData::invoiceItemProperties[i].key);
            }
            else
            {
                for (int l = 0; l < InvoiceDetailData::invoiceDetailItemProperties.size(); ++l)
                {
                    itemNames.push_back(InvoiceDetailData::invoiceDetailItemProperties[i].key);
                }
            }
        }
    }

    for (int i = 0; i < itemNames.size(); ++i)
    {
        output.push_back(itemNames[i]);
    }
}
