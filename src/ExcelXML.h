/****************************************************************************
*MIT License
*
*Copyright (c) 2021 amt.arasan
*
*Permission is hereby granted, free of charge, to any person obtaining a copy
*of this software and associated documentation files (the "Software"), to deal
*in the Software without restriction, including without limitation the rights
*to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
*copies of the Software, and to permit persons to whom the Software is
*furnished to do so, subject to the following conditions:
*
*The above copyright notice and this permission notice shall be included in all
*copies or substantial portions of the Software.
*
*THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
*IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
*FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
*AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
*LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
*OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
*SOFTWARE.
*
*/

#ifndef EXCELXML_H
#define EXCELXML_H

#include <QObject>
#include <QColor>
#include "xmlworkbook.h"

#include<QMap>

class ExcelXMLCellStyle
{
public:
    enum FontUnderline
    {
        FontUnderlineNone,
        FontUnderlineSingle,
        FontUnderlineDouble        
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

    QColor mCellColor = QColor(Qt::white);
    QColor mFontColor = QColor(Qt::black);

    BorderType mBorderType = NoBorder;
    BorderStyle mBorderStyle = BorderNone;
    FontUnderline mFontUnderline = FontUnderlineNone;
};

class LIBABASHARED_EXPORT ExcelXML
{

public:
    explicit ExcelXML();
    virtual ~ExcelXML();

    //To Create Excel Workbook with XML Tags
    bool createWorkBook();

    //To Create Excel Sheet with XML Tags
    //@ istrSheetName - Sheet Name
    bool createWorkSheet(const QString &istrSheetName);

    //To Active Excel Sheet by Sheet Name if Available
    //@ istrSheetName - Sheet Name
    bool activeWorkSheet(const QString &istrSheetName);

    //Add Value to Cell in active Sheet
    //@ iRow - cell Row number
    //@ iColumn - cell column number
    //@ value - cell Value
    bool addValue(int iRow, int iColumn, QVariant &value);

    //Add Value to Cell in active Sheet
    //@ iRow - cell Row number
    //@ iColumn - cell column number
    //@ value - cell Value
    bool setStyleToCell(int iRow, int iColumn, ExcelXMLCellStyle & iStyle);

    //Add List Value
    //@ iStartRow - Row number for add Licst Value
    //@ iColumn - column number for add Licst Value
    //@ value - List value
    bool addListValue(int iStartRow, int iStartColumn, QList<QVariant> &value);

    //Save File as XML
    //@ istrFileNamewithPath - File Name with full Path
    bool save(const QString & istrFileNamewithPath);

private:
    XMLCell* getCell(int iRow, int iColumn);

    XMLRow* getOrCreateRow(int iRow);

private:

    XmlWorkBook * mExcelXMLWorkBook = nullptr;
    XMLWorkSheet * mExcelXMLActiveSheet= nullptr;

    QMap<QString,XMLWorkSheet *> mExcelXMLWorkSheets;

    XMLTable * mExcelXMLTable= nullptr;
    QMap<QString,XMLTable *> mExcelXMLTables;

};

#endif // EXCELXML_H
