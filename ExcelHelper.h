#ifndef EXCELHELPER_H
#define EXCELHELPER_H

#include <QAxObject>
#include <QFile>
#include <QDir>
#include <QDebug>

#define FREE(p) delete p; p = nullptr;

enum class ExcelFormat
{
    xlWorkbookDefault = 51,
    xlOpenXMLWorkbook = 51,
    xlExcel8 = 56,
    xlOpenXMLStrictWorkbook = 61,
};

enum class Alignment
{
    xlCenter = -4108,
    xlTop = -4160,
    xlBottom = -4107,
    xlLeft = -4131,
    xlRight = -4152,
};

enum class NumberFormat
{
    none = 0,
    general,
    currency,
    percentage,
    date,
    time,
    text,
};

class ExcelHelper : public QObject
{
    Q_OBJECT

public:
    ExcelHelper();
    ~ExcelHelper();

    void Kick();
    void SetVisible(bool value);
    void SetCaption(const QString &cap);

    void Create(const QString &fp);
    void Open(const QString &fp);
    void Close();
    void SaveAs(const QString &fp);
    void Save();

    int GetSheetCount();
    void SetCurrentSheet(int index);
    QStringList GetWorkSheetNames();
    QString GetCurrentWorkSheetName();

    void GetUsedRange(int &rowStart, int &colStart, int &rowEnd, int &colEnd);
    void ReadCurrentWorkSheet(QVariant &excelData);
    void WriteCurrentWorkSheet(const QVariant &excelData, const QString &excelDataRange);

    //void SetCellValue(QAxObject *pWorkSheet, int row, int column, const QString &value)
    //{
    //    QAxObject *pRange = pWorkSheet->querySubObject("Cells(int, int)", row, column);
    //    pRange->dynamicCall("Value", value);

    // xlLeft: -4131; xlCenter: -4108; xlRight: -4152
    // xlTop: -4160; xlCenter: -4108; xlBottom: -4107
    //    pRange->dynamicCall("ClearContents()");
    //    QAxObject *pInterior = pRange->querySubObject("Interior");
    //    pInterior->setProperty("Color", QColor(0, 255, 0));
    //    QAxObject *pBorder = cell->querySubObject("Borders");
    //    pBorder->setProperty("Color", QColor(0, 0, 255));
    //    QAxObject *pFont = cell->querySubObject("Font");
    //    pFont->setProperty("Name", QStringLiteral("华文彩云"));
    //    pFont->setProperty("Bold", true);
    //    pFont->setProperty("Size", 20);
    //    pFont->setProperty("Italic", true);
    //    pFont->setProperty("Underline", 2);
    //    pFont->setProperty("Color", QColor(255, 0, 0));
    //    if (row == 1)
    //    {
    //        QAxObject *pFont = pRange->querySubObject("Font");
    //        pFont->setProperty("Bold", true);
    //    }
    //}

    void SetNumberFormat(const QString &format, const QString &range);
    void SetColumnHidden(bool isHidden, const QString &range);
    void SetRowHidden(bool isHidden, const QString &range);

    void SetCellSize(int row, int col, int height, int width);
    void SetCellAlign(int row, int col, Alignment hAlign, Alignment vAlign);
    void SetCellWrap(int row, int col, bool isWrap);
    void SetCellFont();
    void SetCellBorder();
    QVariant GetCell(int row, int col);
    void SetCell(int row, int col, const QVariant &value);

public slots:
    void error(int code, const QString &source, const QString &desc, const QString &help);

private:
    QAxObject *pExcel = nullptr;
    QAxObject *pWorkBooks = nullptr;
    QAxObject *pWorkBook = nullptr;
    QAxObject *pWorkSheets = nullptr;
    QAxObject *pWorkSheet = nullptr;
    QString filePath;
    QString workSheetName;
};

#endif // EXCELHELPER_H
