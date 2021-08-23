#ifndef EXCELXML_H
#define EXCELXML_H

#include <QObject>
#include <QColor>
#include "xmlworkbook.h"

#include<QMap>

class LIBABASHARED_EXPORT ExcelXMLCellStyle
{
public:
    ExcelXMLCellStyle();
    ~ExcelXMLCellStyle();
    enum FontUnderline
    {
        FontUnderlineNone,
        FontUnderlineSingle,
        FontUnderlineDouble
        //FontUnderlineSingleAccounting,
        //FontUnderlineDoubleAccounting
    };

    enum HorizontalAlignment
    {
        AlignHLeft,
        AlignHCenter,
        AlignHRight
    };
    HorizontalAlignment mHAlign = AlignHLeft;

    enum VerticalAlignment
    {
        AlignVTop,
        AlignVCenter,
        AlignVBottom,
    };
    VerticalAlignment mVAlign = AlignVBottom;

    enum BorderStyle
    {
        BorderNone,
        BorderThin,
        BorderMedium,
        BorderDashed,
        BorderDotted,
        BorderThick,
        BorderDouble,
        BorderHair,
        BorderMediumDashed,
        BorderDashDot,
        BorderMediumDashDot,
        BorderDashDotDot,
        BorderMediumDashDotDot,
        BorderSlantDashDot
    };

    enum BorderType
    {
        NoBorder,
        BorderDown,
        BorderUp,
        BorderRight,
        BorderLeft,
        BorderAll
    };

    int mFontSize = 11;
    QString mFontName = "Calibri";
    QString mFontFamily = "Swiss";

    bool mBold = false;
    bool mUnderline = false;
    bool mFontItalic = false;
    bool mWrapText = false;

    QColor mCellColor = QColor(Qt::white);
    QColor mFontColor = QColor(Qt::black);

    BorderType mBorderType = NoBorder;
    BorderStyle mBorderStyle = BorderNone;
    FontUnderline mFontUnderline = FontUnderlineNone;

    uint mMergeCells = 0;
    uint mMergeRows = 0;

    uint mRowHeight = 15;
    int mTextRotate = 0;
};

class LIBABASHARED_EXPORT ExcelXML
{

public:
    explicit ExcelXML();
    virtual ~ExcelXML();

    ///
    /// \brief createWorkBook - Create Excel Workbook with XML Tags
    /// \return - if Success return True or False
    ///
    bool createWorkBook();

    ///
    /// \brief createWorkSheet - Create Excel Sheet with XML Tags
    /// \param istrSheetName   - Sheet name Create on WorkBook
    /// \return - if Success return True or False
    ///
    bool createWorkSheet(const QString &istrSheetName);

    ///
    /// \brief activeWorkSheet - To Active Excel Sheet by Sheet Name if Available
    /// \param istrSheetName - Sheet Name to Activate
    /// \return - if Success return True or False
    ///
    bool activeWorkSheet(const QString &istrSheetName);

    ///
    /// \brief addValue - Add Value to Cell in active Sheet
    /// \param iRow - cell Row number >= 1
    /// \param iColumn - cell column number >= 1 ==> 1->A 2->B 3->C ....
    /// \param value - cell Value
    /// \param iStyle - cell Style
    /// \return - if Success return True or False
    ///
    bool addValue(int iRow, int iColumn, QVariant &value , ExcelXMLCellStyle & iStyle = ExcelXMLCellStyle());

    ///
    /// \brief setStyleToCell - Update Style to the Input Cell
    /// \param iRow    -  cell Row number >= 1
    /// \param iColumn  -  cell column number >= 1 ==> 1->A 2->B 3->C ....
    /// \param iStyle   - cell Style to Update
    /// \return - if Success return True or False
    ///
    bool setStyleToCell(int iRow, int iColumn, ExcelXMLCellStyle & iStyle);

    ///
    /// \brief addListValue - Add List of Value to the Row
    /// \param iStartRow    - Row Number to Add Values >= 1
    /// \param iStartColumn - Column Number to Append the Values >= 1 ==> 1->A 2->B 3->C ....
    /// \param value        - List of Values to Add
    /// \param iStyle       - Style of the cell
    /// \return - if Success return True or False
    ///
    bool addListValue(int iStartRow, int iStartColumn, QList<QVariant> &value, ExcelXMLCellStyle & iStyle = ExcelXMLCellStyle());

    ///
    /// \brief save - Save File as XML in Local Path
    /// \param istrFileNamewithPath - File Name with full Path
    /// \return - if Success return True or False
    ///
    bool save(const QString & istrFileNamewithPath);

private:
    ///
    /// \brief getCell - return Cell in Row
    /// \param iRow  - Row Number  of Cell >= 1
    /// \param iColumn - Colum number of Cell  >= 1 ==> 1->A 2->B 3->C ....
    /// \return - if Success return XMLCell* or nullptr
    ///
    XMLCell* getCell(int iRow, int iColumn);

    ///
    /// \brief getCellbyIndex Get Cell By Input Index (1 ==> 1->A 2->B 3->C ....)
    /// \param iRow  - Row Number  of Cell >= 1
    /// \param iCell  - Colum number of Cell  >= 1 ==> 1->A 2->B 3->C ....
    /// \return if Success return XMLCell* or NULL
    ///
    XMLCell* getCellbyIndex(int iRow, int iCell);

    ///
    /// \brief getOrCreateRow - Create Row or Row Will Return If already Exists
    /// \param iRow - Row Number to Get From the Sheet >= 1
    /// \return if Success return XMLRow* or nullptr
    ///
    XMLRow* getOrCreateRow(int iRow);

private:

    XmlWorkBook * mExcelXMLWorkBook = nullptr;
    XMLWorkSheet * mExcelXMLActiveSheet= nullptr;

    QMap<QString,XMLWorkSheet *> mExcelXMLWorkSheets;

    XMLTable * mExcelXMLTable= nullptr;
    QMap<QString,XMLTable *> mExcelXMLTables;

};

#endif // EXCELXML_H
