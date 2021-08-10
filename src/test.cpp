#include "ExcelXML.h"
#include <iostream>


int main(int argc, char *argv[])
{
  ExcelXML xmlworkbook;
    if(xmlworkbook.createWorkBook())
    {
        if(xmlworkbook.createWorkSheet("Demo"))
        {
            xmlworkbook.addValue(5,5,QVariant("Demo1"));
            xmlworkbook.addValue(5,6,QVariant("Demo2"));
            xmlworkbook.addValue(5,7,QVariant("Demo3"));
            xmlworkbook.addValue(6,5,QVariant("Demo4"));
            xmlworkbook.addValue(7,5,QVariant("Demo5"));
            xmlworkbook.addValue(8,5,QVariant("Demo6"));

            xmlworkbook.addValue(5,5,QVariant(int(256)));
            xmlworkbook.addValue(6,5,QVariant(float(256.204)));

            ExcelXMLCellStyle iStyle;
            iStyle.mBorderStyle = ExcelXMLCellStyle::BorderMedium;
            iStyle.mBorderType = ExcelXMLCellStyle::BorderAll;
            iStyle.mBold = true;
            iStyle.mFontColor = Qt::green;
            iStyle.mCellColor = Qt::yellow;

            xmlworkbook.setStyleToCell(5,5,iStyle);

            iStyle.mBold = false;
            iStyle.mUnderline = true;
            iStyle.mFontColor = Qt::red;
            iStyle.mCellColor = Qt::black;
            xmlworkbook.setStyleToCell(6,5,iStyle);
        }

        if(xmlworkbook.createWorkSheet("Demo1"))
        {
            xmlworkbook.addValue(15,5,QVariant("Demo1"));
            xmlworkbook.addValue(15,6,QVariant("Demo2"));
            xmlworkbook.addValue(15,7,QVariant("Demo3"));
            xmlworkbook.addValue(16,5,QVariant("Demo4"));
            xmlworkbook.addValue(17,5,QVariant("Demo5"));
            xmlworkbook.addValue(18,5,QVariant("Demo6"));

            xmlworkbook.addValue(15,5,QVariant(int(256)));
            xmlworkbook.addValue(16,5,QVariant(float(256.204)));

            ExcelXMLCellStyle iStyle;
            iStyle.mBorderStyle = ExcelXMLCellStyle::BorderMedium;
            iStyle.mBorderType = ExcelXMLCellStyle::BorderAll;
            iStyle.mBold = true;
            iStyle.mFontColor = Qt::green;
            iStyle.mCellColor = Qt::yellow;

            xmlworkbook.setStyleToCell(15,5,iStyle);

            iStyle.mBold = false;
            iStyle.mUnderline = true;
            iStyle.mFontColor = Qt::red;
            iStyle.mCellColor = Qt::black;
            iStyle.mHAlign = ExcelXMLCellStyle::AlignHCenter;
            iStyle.mVAlign = ExcelXMLCellStyle::AlignVCenter;
            xmlworkbook.setStyleToCell(16,5,iStyle);
        }
    }
	xmlworkbook.save("C:\\Tamil\\demo.xml");
	
	return 0;

}