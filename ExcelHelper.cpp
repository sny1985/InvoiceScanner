#include "ExcelHelper.h"

ExcelHelper::ExcelHelper() : QObject()
{
    pExcel = new QAxObject(this);
    pExcel->setControl("Excel.Application");
    if (pExcel->isNull())
    {
        pExcel->setControl("ket.Application");
    }

    if (!pExcel->isNull())
    {
        pExcel->setProperty("Visible", false);
        pExcel->setProperty("DisplayAlerts", false);

        pWorkBooks = pExcel->querySubObject("Workbooks"); // Returns a Workbooks collection that represents all the open workbooks. Read-only.
    }
}

ExcelHelper::~ExcelHelper()
{
    Close();

    if (pExcel != nullptr && !pExcel->isNull())
    {
        pExcel->dynamicCall("Quit()");
    }
    FREE(pExcel)

    disconnect();
}

void ExcelHelper::Kick()
{
    if (pExcel != nullptr && !pExcel->isNull())
    {
        pExcel->setProperty("Visible", true);
    }
    Close();
}

void ExcelHelper::SetVisible(bool value)
{
    if (pExcel != nullptr && !pExcel->isNull())
    {
        pExcel->setProperty("Visible", value);
    }
}

void ExcelHelper::SetCaption(const QString &cap)
{
    if (pExcel != nullptr && !pExcel->isNull())
    {
        pExcel->setProperty("Caption", cap);
    }
}

void ExcelHelper::Create(const QString &fp)
{
    if (pWorkBooks != nullptr && ! pWorkBooks->isNull())
    {
        pWorkBooks->dynamicCall("Add()"); // Creates a new workbook. The new workbook becomes the active workbook.
        pWorkBook = pExcel->querySubObject("ActiveWorkBook"); // Returns a Workbook object that represents the workbook in the active window.
        pWorkSheets = pWorkBook->querySubObject("WorkSheets"); // Returns a Sheets collection that represents all the worksheets in the specified workbook. Read-only Sheets object.
        pWorkSheet = pWorkBook->querySubObject("ActiveSheet"); // Returns a Worksheet object that represents the active sheet (the sheet on top) in the active workbook or specified workbook.
        bool rs = pWorkSheet != nullptr && !pWorkSheet->isNull();
        workSheetName = rs ? pWorkSheet->property("Name").toString() : "";
        filePath = fp;
    }
}

void ExcelHelper::Open(const QString &fp)
{
    Close();

    if (pWorkBooks != nullptr && ! pWorkBooks->isNull() )
    {
        pWorkBook = pWorkBooks->querySubObject("Open(const QString &)", fp); // Opens a workbook.
        if (pWorkBook != nullptr && ! pWorkBook->isNull())
        {
            pWorkSheets = pWorkBook->querySubObject("WorkSheets");
            pWorkSheet = pWorkBook->querySubObject("ActiveSheet");
            bool rs = pWorkSheet != nullptr && !pWorkSheet->isNull();
            workSheetName = rs ? pWorkSheet->property("Name").toString() : "";
            filePath = fp;
        }
    }
}

void ExcelHelper::Close()
{
    FREE(pWorkSheet)
    FREE(pWorkSheets)
    if (pWorkBook != nullptr && ! pWorkBook->isNull())
    {
        pWorkBook->dynamicCall("Close(Boolean)", false); // Closes the object.
    }
    FREE(pWorkBook)
//        FREE(pWorkBooks)
    filePath.clear();
    workSheetName.clear();
}

void ExcelHelper::SaveAs(const QString &fp)
{
    if (pWorkBook != nullptr && ! pWorkBook->isNull() )
    {
        filePath = fp;
        pWorkBook->dynamicCall("SaveAs(const QString &, int, const QString &, const QString &, bool, bool)",
                               QDir::toNativeSeparators(filePath), static_cast<int>(ExcelFormat::xlWorkbookDefault), QString(""), QString(""), false, false); // Saves changes to the workbook in a different file.
    }
}

void ExcelHelper::Save()
{
    if (!filePath.isEmpty())
    {
        SaveAs(filePath);
    }
}

int ExcelHelper::GetSheetCount()
{
    if (pWorkSheets != nullptr && !pWorkSheets->isNull())
    {
        return pWorkSheets->property("Count").toInt();
    }
    return 0;
}

void ExcelHelper::SetCurrentSheet(int index)
{
    if (pWorkSheets != nullptr && !pWorkSheets->isNull())
    {
        FREE(pWorkSheet)
        pWorkSheet = pWorkSheets->querySubObject("Item(int)", index); // Returns a single object from a collection.
        bool rs = pWorkSheet != nullptr && ! pWorkSheet->isNull();
        if (rs)
        {
            pWorkSheet->dynamicCall("Activate()"); // Makes the current sheet the active sheet.
        }
        workSheetName = rs ? pWorkSheet->property("Name").toString() : "";
    }
}

QStringList ExcelHelper::GetWorkSheetNames()
{
    QStringList rs;
    if (pWorkSheets != nullptr && !pWorkSheets->isNull())
    {
        int sheetCount = pWorkSheets->property("Count").toInt();
        for (int i = 1; i <= sheetCount; ++i)
        {
            QAxObject *pSheet = pWorkSheets->querySubObject("Item(int)", i);
            if (pSheet != nullptr && !pSheet->isNull())
            {
                rs.append(pSheet->property("Name").toString());
                delete pSheet;
            }
        }
    }
    return rs;
}

QString ExcelHelper::GetCurrentWorkSheetName()
{
    return workSheetName;
}

void ExcelHelper::GetUsedRange(int &rowStart, int &colStart, int &rowEnd, int &colEnd)
{
    if (pWorkSheet != nullptr && !pWorkSheet->isNull())
    {
        QAxObject *pRange = pWorkSheet->querySubObject("UsedRange");
        rowStart = pRange->property("Row").toInt();
        colStart = pRange->property("Column").toInt();
        rowEnd = pRange->querySubObject("Rows")->property("Count").toInt();
        colEnd = pRange->querySubObject("Columns")->property("Count").toInt();
        delete pRange;
    }
}

void ExcelHelper::ReadCurrentWorkSheet(QVariant &excelData)
{
    if (pWorkSheet && !pWorkSheet->isNull())
    {
        QAxObject *pUsedRange = pWorkSheet->querySubObject("UsedRange"); // Returns a Range object that represents the used range on the specified worksheet. Read-only.
        if (pUsedRange && !pUsedRange->isNull())
        {
            excelData = pUsedRange->property("Value");
        }
        delete pUsedRange;
    }
}

void ExcelHelper::WriteCurrentWorkSheet(const QVariant &excelData, const QString &excelDataRange)
{
    if (pWorkSheet && !pWorkSheet->isNull())
    {
        QAxObject *pRange = pWorkSheet->querySubObject("Range(const QString &, const QString &)", excelDataRange); // Returns a Range object that represents a cell or a range of cells.
        if (pRange && !pRange->isNull())
        {
            pRange->setProperty("Value", excelData);
        }
        delete pRange;
    }
}

void ExcelHelper::SetNumberFormat(const QString &format, const QString &range)
{
    if (pWorkSheet != nullptr && !pWorkSheet->isNull())
    {
        QAxObject *pRange = pWorkSheet->querySubObject("Range(const QString &, const QString &)", range);
        if (pRange && !pRange->isNull())
        {
            pRange->setProperty("NumberFormatLocal", format);
        }
        delete pRange;
    }
}

void ExcelHelper::SetColumnHidden(bool isHidden, const QString &range)
{
    if (pWorkSheet != nullptr && !pWorkSheet->isNull())
    {
        QAxObject *pRange = pWorkSheet->querySubObject("Range(const QString &, const QString &)", range);
        if (pRange && !pRange->isNull())
        {
            QAxObject *pCol = pRange->querySubObject("EntireColumn");
            if (pCol && !pCol->isNull())
            {
                pCol->setProperty("Hidden", isHidden);
            }
            delete pCol;
        }
        delete pRange;
    }
}

void ExcelHelper::SetRowHidden(bool isHidden, const QString &range)
{
    if (pWorkSheet != nullptr && !pWorkSheet->isNull())
    {
        QAxObject *pRange = pWorkSheet->querySubObject("Range(const QString &, const QString &)", range);
        if (pRange && !pRange->isNull())
        {
            QAxObject *pRow = pRange->querySubObject("EntireRow");
            if (pRow && !pRow->isNull())
            {
                pRow->setProperty("Hidden", isHidden);
            }
            delete pRow;
        }
        delete pRange;
    }
}

void ExcelHelper::SetCellSize(int row, int col, int height, int width)
{
    if (pWorkSheet != nullptr && !pWorkSheet->isNull())
    {
        QAxObject *pRange = pWorkSheet->querySubObject("Cells(int, int)", row, col);
        pRange->setProperty("RowHeight", height);
        pRange->setProperty("ColumnWidth", width);
        delete pRange;
    }
}

void ExcelHelper::SetCellAlign(int row, int col, Alignment hAlign, Alignment vAlign)
{
    if (pWorkSheet != nullptr && !pWorkSheet->isNull())
    {
        QAxObject *pRange = pWorkSheet->querySubObject("Cells(int, int)", row, col);
        pRange->setProperty("HorizontalAlignment", static_cast<int>(hAlign));
        pRange->setProperty("VerticalAlignment", static_cast<int>(vAlign));
        delete pRange;
    }
}

void ExcelHelper::SetCellWrap(int row, int col, bool isWrap)
{
    if (pWorkSheet != nullptr && !pWorkSheet->isNull())
    {
        QAxObject *pRange = pWorkSheet->querySubObject("Cells(int, int)", row, col);
        pRange->setProperty("WrapText", isWrap);
        delete pRange;
    }
}

void ExcelHelper::SetCellFont()
{

}

void ExcelHelper::SetCellBorder()
{

}

QVariant ExcelHelper::GetCell(int row, int col)
{
    QVariant rs;
    if (pWorkSheet != nullptr && !pWorkSheet->isNull())
    {
        QAxObject *pRange = pWorkSheet->querySubObject("Cells(int, int)", row, col);
        rs = pRange->dynamicCall("Value()");
        delete pRange;
    }
    return rs;
}

void ExcelHelper::SetCell(int row, int col, const QVariant &value)
{
    if (pWorkSheet != nullptr && !pWorkSheet->isNull())
    {
        QAxObject *pRange = pWorkSheet->querySubObject("Cells(int, int)", row, col);
        pRange->setProperty("Value", value);
        delete pRange;
    }
}

// Usage:
//        pObj->blockSignals(false);
//        connect(pObj, SIGNAL(exception(int, const QString &, const QString &, const QString &)), this, SLOT(error(int, const QString &, const QString &, const QString &)));
void ExcelHelper::error(int code, const QString &source, const QString &desc, const QString &help)
{
    qDebug() << code;
    qDebug() << source;
    qDebug() << desc;
    qDebug() << help;
}
