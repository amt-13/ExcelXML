#include "ExcelXML.h"
#include <QFileInfo>
#include <QVariant>

#include <QDebug>

ExcelXML::ExcelXML()
{
    mExcelXMLWorkSheets.clear();

    mExcelXMLWorkBook = nullptr;
    mExcelXMLActiveSheet= nullptr;
}

ExcelXML::~ExcelXML()
{
    mExcelXMLWorkSheets.clear();
    mExcelXMLWorkBook = nullptr;
    mExcelXMLActiveSheet= nullptr;
}

bool ExcelXML::createWorkBook()
{
    mExcelXMLWorkBook = new XmlWorkBook();

    XmlStyle defaultStyle; // Default style.
    mExcelXMLWorkBook->Styles.push_back(defaultStyle);

    return true;
}

bool ExcelXML::createWorkSheet(const QString &istrSheetName)
{
    mExcelXMLActiveSheet = new XMLWorkSheet();
    mExcelXMLActiveSheet->Name = istrSheetName.toStdString();

    mExcelXMLWorkSheets[istrSheetName]  = mExcelXMLActiveSheet;

    mExcelXMLTable = new XMLTable();
    mExcelXMLTables[istrSheetName]  = mExcelXMLTable;
    return true;
}

bool ExcelXML::activeWorkSheet(const QString &istrSheetName)
{
    mExcelXMLActiveSheet = mExcelXMLWorkSheets[istrSheetName];
    return true;
}

bool ExcelXML::addValue(int iRow, int iColumn, QVariant &value)
{
    if(iRow < 1 || iColumn < 1)
        return false;

    XMLRow * pRow = getOrCreateRow(iRow);
    qDebug() << "pRow->Cells.size() " <<  pRow->Cells.size();

    if(pRow->Cells.size() >= iColumn)
    {
        XMLCell *pxmlCell = &(pRow->Cells[iColumn-1]);
        if(pxmlCell)
        {
            if(value.type() == QVariant::String)
            {
                pxmlCell->Type = "String";
                pxmlCell->Value = value.toString().toStdString();

            }
            else if(value.type() == QVariant::Int)
            {
                pxmlCell->Type = "Number";
                pxmlCell->Value = QString::number(value.toInt()).toStdString();
            }
            else if(value.type() == QVariant::Double )
            {
                pxmlCell->Type = "Number";
                pxmlCell->Value = QString::number(value.toDouble()).toStdString();
            }
            else if(value.type() == QVariant::LongLong)
            {
                pxmlCell->Type = "Number";
                pxmlCell->Value = QString::number(value.toLongLong()).toStdString();
            }
            else if(value.type() == QVariant::UInt )
            {
                pxmlCell->Type = "Number";
                pxmlCell->Value = QString::number(value.toUInt()).toStdString();
            }
            else if(value.type() == QVariant::ULongLong)
            {
                pxmlCell->Type = "Number";
                pxmlCell->Value = QString::number(value.toULongLong()).toStdString();
            }
            else
            {
                pxmlCell->Type = "Number";
                pxmlCell->Value = QString::number(value.toFloat()).toStdString();
            }
            return true;
        }
    }else
    {
        for (size_t col = pRow->Cells.size()+1; col <= iColumn; ++col)
        {
            XMLCell xmlCell;
            qDebug() << "col " <<  col;
            qDebug() << "iColumn " <<  iColumn;
            if(col == iColumn)
            {
                if(value.type() == QVariant::String)
                {
                    xmlCell.Type = "String";
                    xmlCell.Value = value.toString().toStdString();
                }
                else if(value.type() == QVariant::Int)
                {
                    xmlCell.Type = "Number";
                    xmlCell.Value = QString::number(value.toInt()).toStdString();
                }
                else if(value.type() == QVariant::Double )
                {
                    xmlCell.Type = "Number";
                    xmlCell.Value = QString::number(value.toDouble()).toStdString();
                }
                else if(value.type() == QVariant::LongLong)
                {
                    xmlCell.Type = "Number";
                    xmlCell.Value = QString::number(value.toLongLong()).toStdString();
                }
                else if(value.type() == QVariant::UInt )
                {
                    xmlCell.Type = "Number";
                    xmlCell.Value = QString::number(value.toUInt()).toStdString();
                }
                else if(value.type() == QVariant::ULongLong)
                {
                    xmlCell.Type = "Number";
                    xmlCell.Value = QString::number(value.toULongLong()).toStdString();
                }
                else
                {
                    xmlCell.Type = "Number";
                    xmlCell.Value = QString::number(value.toFloat()).toStdString();
                }
            }
            else
            {
                xmlCell.Type = "String";
                xmlCell.Value = "";
            }

            xmlCell.Style = mExcelXMLWorkBook->Styles[0].ID;
            pRow->Cells.push_back(xmlCell);
        }
    }

    return true;
}

bool ExcelXML::setStyleToCell(int iRow, int iColumn, ExcelXMLCellStyle &iStyle)
{
    if(iRow < 1 || iColumn < 1)
        return false;

    XMLRow * pRow = getOrCreateRow(iRow);
    if(pRow->Cells.size() >= iColumn)
    {
        XMLCell *pxmlCell = &(pRow->Cells[iColumn-1]);
        if(pxmlCell)
        {
            // ************ Setting Style Starts ************ //
            QString strStyleID = "style_" + QString::number(iRow)+ "_" + QString::number(iColumn);
            XmlStyle style(strStyleID.toStdString());

            bool ok;
            QString fcolor = iStyle.mFontColor.name();
            fcolor.replace("#","");
            long font_color = fcolor.toLong(&ok,16);

            style.setFont(iStyle.mFontName.toStdString(), iStyle.mFontFamily.toStdString(), iStyle.mFontSize, font_color);

            if(iStyle.mBold)
                style.appendFontBold();

            if(iStyle.mUnderline)
            {
                if(iStyle.mFontUnderline == ExcelXMLCellStyle::FontUnderlineDouble)
                    style.appendFontDoubleUnderline();
                else
                    style.appendFontSingleUnderline();
            }

            if(iStyle.mCellColor.isValid())
            {
                QString color = iStyle.mCellColor.name();
                color.replace("#","");
                long cell_color = color.toLong(&ok,16);
                style.setInterior(cell_color);
            }

            if(iStyle.mBorderType ==  ExcelXMLCellStyle::BorderAll)
                style.setBorders(1,1,1,1);
            else if(iStyle.mBorderType ==  ExcelXMLCellStyle::BorderDown)
                style.setBorders(0,1,0,0);
            else if(iStyle.mBorderType ==  ExcelXMLCellStyle::BorderUp)
                style.setBorders(1,0,0,0);
            else if(iStyle.mBorderType ==  ExcelXMLCellStyle::BorderLeft)
                style.setBorders(0,0,1,0);
            else if(iStyle.mBorderType ==  ExcelXMLCellStyle::BorderRight)
                style.setBorders(0,0,0,1);

            style.setAlignment((XmlStyle::HorizontalAlignment)iStyle.mHAlign,(XmlStyle::VerticalAlignment)iStyle.mVAlign);


            // ************ Setting Style Ends************ //


            // Compare Existing Style to new style to Avoid Duplicate style tag Creation
            for (size_t i = 0; i < mExcelXMLWorkBook->Styles.size(); i++)
            {
               if(style == mExcelXMLWorkBook->Styles[i])
               {
                   pxmlCell->Style = mExcelXMLWorkBook->Styles[i].ID;
                   return true;
               }
               else if(strStyleID.compare(QString(mExcelXMLWorkBook->Styles[i].ID.c_str())) == 0)
               {
                   pxmlCell->Style = mExcelXMLWorkBook->Styles[i].ID;
                   return true;
               }
            }

            // Push the Style to XML
            mExcelXMLWorkBook->Styles.push_back(style);
            pxmlCell->Style = style.ID;

            return true;
        }
    }

    return false;
}

bool ExcelXML::addListValue(int iStartRow, int iStartColumn, QList<QVariant> &value)
{

    return true;
}

XMLRow* ExcelXML::getOrCreateRow(int iRow)
{
    std::vector<XMLRow> listRows = mExcelXMLTable->Rows;
    if(listRows.size() < iRow)
    {
        for (size_t row = listRows.size(); row < iRow; ++row)
        {
            mExcelXMLTable->insertRow(row,NULL);
        }
    }

    return &(mExcelXMLTable->Rows[iRow-1]);
}

XMLCell* ExcelXML::getCell(int iRow, int iColumn)
{
    std::vector<XMLRow> listRows = mExcelXMLTable->Rows;
    for (size_t row = 0; row < listRows.size(); ++row)
    {
        XMLRow xmlRow = listRows[row];
        std::vector<XMLCell> listCells = xmlRow.Cells;
        for (size_t col = 0; col < listCells.size(); ++col)
        {
            if(row == iRow && col == iColumn)
            {
                return &(listCells[col]);
            }
        }
    }

    return nullptr;
}

bool ExcelXML::save(const QString &istrFileNamewithPath)
{
    if(mExcelXMLWorkBook == nullptr)
    {
        return false;
    }

    if(istrFileNamewithPath.length() <= 0)
    {
        return false;
    }

    QFileInfo file(istrFileNamewithPath);
    if(file.suffix().compare("xml",Qt::CaseInsensitive) != 0)
    {
        return false;
    }

    QFile outFile(istrFileNamewithPath);
    if(outFile.open(QFile::WriteOnly) == false)
    {
        return false;
    }

    QList <XMLWorkSheet*> listSheets = mExcelXMLWorkSheets.values();
    if(listSheets.count() <= 0)
    {
        return false;
    }

    QList<QString> keys = mExcelXMLWorkSheets.keys();
    for(int i = 0 ; i < listSheets.count(); ++i)
    {
        listSheets[i]->Tables.push_back(*(mExcelXMLTables[keys[i]]));
        mExcelXMLWorkBook->Sheets.push_back(*listSheets[i]);
    }

    outFile.write(mExcelXMLWorkBook->xml().c_str());
    outFile.close();

    return true;
}
