#include "xmlworkbook.h"

#include <sstream>
#include <iomanip>

/////////////////////////////////////////////////////////////////////////////
// XmlStyle class
/////////////////////////////////////////////////////////////////////////////

/// Construct as default style.
XmlStyle::XmlStyle(const XMLSTR & id) :
    ID(id),
    Name(id == "Default" ? "Normal" :""),
    Alignment("ss:Vertical=\"Bottom\""),
    Font("ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Size=\"11\" ss:Color=\"#000000\""),
    Borders(""),
    Interior(""),
    NumberFormat(""),
    Protection("")
{   
}

/// Destructor.
XmlStyle::~XmlStyle()
{
}

/// Output as XML.
const XMLSTR XmlStyle::xml()
{
    std::stringstream rv; 
    XMLSTR eol = "\n";

    rv << "  <Style ss:ID=\"" << ID << "\"";
    if (!Name.empty())         rv << " ss:Name=\"" << Name << "\"";
    rv << ">" << eol;
    
    if (!Alignment.empty())    rv << "   <Alignment " << Alignment << "/>" << eol;
    if (!Borders.empty())      rv << "   <Borders>" << eol << Borders << "   </Borders>" << eol;
    if (!Font.empty())         rv << "   <Font " << Font << "/>" << eol;
    if (!Interior.empty())     rv << "   <Interior " << Interior << "/>" << eol;
    if (!NumberFormat.empty()) rv << "   <NumberFormat " << NumberFormat << "/>" << eol;
    if (!Protection.empty())   rv << "   <Protection " << Protection << "/>" << eol;
    rv << "  </Style>" << eol;

    return rv.str();
}

const XMLSTR XmlStyle::xmlwithoutIDName()
{
    std::stringstream rv;
    XMLSTR eol = "\n";

    rv << "  <Style> " << eol;
    if (!Alignment.empty())    rv << "   <Alignment " << Alignment << "/>" << eol;
    if (!Borders.empty())      rv << "   <Borders>" << eol << Borders << "   </Borders>" << eol;
    if (!Font.empty())         rv << "   <Font " << Font << "/>" << eol;
    if (!Interior.empty())     rv << "   <Interior " << Interior << "/>" << eol;
    if (!NumberFormat.empty()) rv << "   <NumberFormat " << NumberFormat << "/>" << eol;
    if (!Protection.empty())   rv << "   <Protection " << Protection << "/>" << eol;
    rv << "  </Style>" << eol;

    return rv.str();
}

/// Set the font string.
void XmlStyle::setFont
    (
        const XMLSTR & fontName,
        const XMLSTR & fontFamily,
        float size,
        long color
    )
{
    std::stringstream s;
    s   << "ss:FontName=\"" << fontName 
        << "\" x:Family=\"" << fontFamily 
        << "\" ss:Size=\"" << size 
        << "\" ss:Color=\"#" << std::uppercase 
        << std::setfill('0') << std::setw(6) 
        << std::hex <<  color << "\"";  

    Font = s.str();
}

/// Set the interior string.
void XmlStyle::setInterior
    (
        long color,
        const XMLSTR & pattern
    )
{
    std::stringstream s;
    s   << "ss:Pattern=\"" << pattern 
        << "\" ss:Color=\"#" << std::uppercase 
        << std::setfill('0') << std::setw(6) 
        << std::hex <<  color << "\"";  
    Interior = s.str();
}

/*
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
<Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/> */
void XmlStyle::setBorders
    (
        size_t topWeight,
        size_t bottomWeight,
        size_t leftWeight,
        size_t rightWeight
    )
{
    std::stringstream s;
    XMLSTR eol = "\n";
    if (topWeight != 0)     s << "    <Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"" << topWeight << "\"/>" + eol;
    if (bottomWeight != 0)  s << "    <Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"" << bottomWeight << "\"/>" + eol;
    if (leftWeight != 0)    s << "    <Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"" << leftWeight << "\"/>" + eol;
    if (rightWeight != 0)   s << "    <Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"" << rightWeight << "\"/>" + eol;


    Borders = s.str();
}

void XmlStyle::setAlignment(XmlStyle::HorizontalAlignment hAlign,
                            XmlStyle::VerticalAlignment vAlign, bool mWrapText, int mTextRotate)
{
    std::stringstream s;

    if (hAlign == AlignHLeft)
        s << "ss:Horizontal=\"Left\"";
    else if (hAlign == AlignHCenter)
        s << "ss:Horizontal=\"Center\"";
    else if (hAlign == AlignHRight)
        s << "ss:Horizontal=\"Right\"";

    s <<" ";

    if (vAlign == AlignVBottom)
        s << "ss:Vertical=\"Bottom\"";
    else if (vAlign == AlignVCenter)
        s << "ss:Vertical=\"Center\"";
    else if (vAlign == AlignVTop)
        s << "ss:Vertical=\"Top\"";

    if(mWrapText)
        s << " ss:WrapText=\"1\"";

    if(mTextRotate > 0)
        s << " ss:Rotate=\"" << mTextRotate << "\"";

    Alignment = s.str();
}

bool XmlStyle::operator== ( XmlStyle& s2)
{
    return (this->xmlwithoutIDName() == s2.xmlwithoutIDName() );
}

bool XmlStyle::operator!= ( XmlStyle& s2)
{
    return !(this->xmlwithoutIDName() == s2.xmlwithoutIDName() );
}



/////////////////////////////////////////////////////////////////////////////
// XMLCell class
/////////////////////////////////////////////////////////////////////////////

/// Constructor.
XMLCell::XMLCell() :
    Style(""),
    Type("Number"),
    Value(""),
    Formula(""),
    Index(0)
{
}

/// Constructor.
XMLCell::~XMLCell()
{
}

/// Output as XML.
const XMLSTR XMLCell::xml()
{
    std::stringstream rv; 
    XMLSTR eol = "\n";
     
    if (Index == 0)
    { 
        rv << "    <Cell";
    }
    else
    {
        rv << "    <Cell ss:Index=\"" << Index << "\"";
    }

    if (mergeCols > 0)     rv << " ss:MergeAcross=\"" << mergeCols << "\"";
    if (mergeRows > 0)     rv << " ss:MergeDown=\"" << mergeRows << "\"";
    if (!Style.empty())     rv << " ss:StyleID=\"" << Style << "\"";
    if (!Formula.empty())   rv << " ss:Formula=\"=" << Formula << "\"";
    rv << "><Data ss:Type=\"" << Type << "\">" << Value << "</Data></Cell>" << eol;    
    return rv.str();
}



/////////////////////////////////////////////////////////////////////////////
// XMLRow class
/////////////////////////////////////////////////////////////////////////////

/// Constructor.
XMLRow::XMLRow() :
    Cells()
{
}

/// Constructor.
XMLRow::~XMLRow()
{
}

/// Output as XML.
const XMLSTR XMLRow::xml()
{
    std::stringstream rv; 
    XMLSTR eol = "\n";

    rv <<  "   <Row ss:AutoFitHeight=\"0\">" << eol;
    for (size_t i = 0; i < Cells.size(); i++)
    {
        rv << Cells.at(i).xml();
    }
    rv << "   </Row>" << eol;
   
    return rv.str();
}

/////////////////////////////////////////////////////////////////////////////
// XMLTable class
/////////////////////////////////////////////////////////////////////////////

/// Constructor.
XMLTable::XMLTable() : 
    RowHeight(15), 
    Rows()
{
}
 
/// Constructor.
XMLTable::~XMLTable()
{
}

/// Output as XML.
const XMLSTR XMLTable::xml()
{
    std::stringstream rv; 
    XMLSTR eol = "\n";

    rv  << "  <Table ss:DefaultRowHeight=\""<< RowHeight <<"\">" << eol;
    for (size_t i = 0; i < Rows.size(); i++)
    {
        rv << Rows.at(i).xml();
    }
    rv << "  </Table>" << eol;

    return rv.str();
}

/// Shift a row up or down.
void XMLTable::insertRow
    (
        size_t whichRow,        ///< Which row.
        XMLRow * row            ///< Row or null to insert a blank row.
    )
{
    std::vector<XMLRow>::iterator it;
    it = Rows.begin();
    if (row == NULL)
    { 
        XMLRow r;
        Rows.insert(it + whichRow, r);
    }
    else
    {
        Rows.insert(it + whichRow, *row);
    }
}    

/////////////////////////////////////////////////////////////////////////////
// XMLWorkSheet class
/////////////////////////////////////////////////////////////////////////////

/// Constructor.
XMLWorkSheet::XMLWorkSheet() :
    Name(""),
    Tables()
{
}

/// Constructor.
XMLWorkSheet::~XMLWorkSheet()
{
}

/// Output as XML.
const XMLSTR XMLWorkSheet::xml()
{
    std::stringstream rv; 
    XMLSTR eol = "\n";

    rv << " <Worksheet ss:Name=\"" << Name << "\">" << eol;

    // Handle the tables.
    for (size_t i = 0; i < Tables.size(); i++)
    {
        rv << Tables.at(i).xml();
    }

    rv << "  <WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\">" << eol;
    rv << "   <PageSetup>" << eol;
    rv << "    <Header x:Margin=\"0.3\"/>" << eol;
    rv << "    <Footer x:Margin=\"0.3\"/>" << eol;
    rv << "    <PageMargins x:Bottom=\"0.75\" x:Left=\"0.7\" x:Right=\"0.7\" x:Top=\"0.75\"/>" << eol;
    rv << "   </PageSetup>" << eol;
    rv << "   <Print>" << eol;
    rv << "    <ValidPrinterInfo/>" << eol;
    rv << "    <HorizontalResolution>-1</HorizontalResolution>" << eol;
    rv << "    <VerticalResolution>-1</VerticalResolution>" << eol;
    rv << "   </Print>" << eol;
    rv << "   <Panes>" << eol;
    rv << "    <Pane>" << eol;
    rv << "     <Number>3</Number>" << eol;
    rv << "     <ActiveRow>5</ActiveRow>" << eol;
    rv << "     <ActiveCol>2</ActiveCol>" << eol;
    rv << "    </Pane>" << eol;
    rv << "   </Panes>" << eol;
    rv << "   <ProtectObjects>False</ProtectObjects>" << eol;
    rv << "   <ProtectScenarios>False</ProtectScenarios>" << eol;
    rv << "  </WorksheetOptions>" << eol;
    rv << " </Worksheet>" << eol;
    
    return rv.str();      
}

void XMLWorkSheet::addTable
    (
        XMLSTR * colNames,     ///< Column names of size cols or null.
        XMLSTR * colNameStyle, ///< Style of column names.
        XMLSTR * dataStyle,    ///< Data style name.
        size_t rows,                ///< Number of rows.
        size_t cols,                ///< Number of columns.
        double * pData,             ///< rows*cols data.
        bool asRows,                ///< True means data is [r0,c0] [r0,c1]... [r0,cn],[r1,c0] ... else by column.
        size_t startRow,            ///< Start Row offset.
        size_t startCol             ///< Start Column offset.
    ) 
{
    // Declare the table.
    XMLTable t;

    // Add the dumy rows.
    for (size_t i = 0; i <startRow; i++)
    {
        XMLRow r;
        t.Rows.push_back(r);
    }

    // Add column headers.
    if (colNames != NULL)
    {
        XMLRow r;
        for (size_t i = 0; i < cols; i++)
        {
            XMLCell c;
            c.Style = colNameStyle != NULL ? *colNameStyle : XMLSTR("");
            c.Type = XMLSTR("String");
            c.Value = colNames[i];
            c.Index = startCol == 0 ? 0 : 1 + startCol + i;
            r.Cells.push_back(c);
        }
        t.Rows.push_back(r);
    }

    // Always form as rows, with columns in spite of indexing.
    for (size_t rr = 0; rr <rows; rr++)
    {
        XMLRow r;
        for (size_t cc = 0; cc < cols; cc++)
        {
            XMLCell c;
            std::stringstream strStream;
            int index = asRows ? (rr * cols) + cc : (cc * rows) + rr;
            strStream << ((double)pData[index]);
            c.Value = strStream.str();
            c.Style = dataStyle != NULL ? *dataStyle : XMLSTR("");
            c.Type = XMLSTR("Number");
            c.Value = strStream.str();
            c.Index = startCol == 0 ? 0 : 1 + startCol + cc;
            r.Cells.push_back(c);
        }

        // Push back row.
        t.Rows.push_back(r);
    }

    // Push back table.
    Tables.push_back(t);
}


/////////////////////////////////////////////////////////////////////////////
// XmlWorkBook class
/////////////////////////////////////////////////////////////////////////////

/// Constructor.
XmlWorkBook::XmlWorkBook() :
    Author(""),
    LastAuthor(""),
    ActiveSheet(1),
    WindowWidth(1000),
    WindowHeight(1000)
{
}

/// Destructor.
XmlWorkBook::~XmlWorkBook()
{
}

/// Header tags, styles, worksheets, endtags.
const XMLSTR XmlWorkBook::xml()
{
    std::stringstream rv;
    rv << headerTags() << stylesTags() << workSheets() << endTags();
    return rv.str();
}

/// Output the xml, workbook, document properties, etc. tags.
const XMLSTR XmlWorkBook::headerTags()
{
    std::stringstream rv; 
    XMLSTR eol = "\n";
    rv 
        << "<?xml version=\"1.0\"?>" << eol 
        << "<?mso-application progid=\"Excel.Sheet\"?>" << eol 
        << "<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"" << eol 
        << " xmlns:o=\"urn:schemas-microsoft-com:office:office\"" << eol 
        << " xmlns:x=\"urn:schemas-microsoft-com:office:excel\"" << eol 
        << " xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"" << eol 
        << " xmlns:html=\"http://www.w3.org/TR/REC-html40\">" << eol 
        << " <DocumentProperties xmlns=\"urn:schemas-microsoft-com:office:office\">" << eol 
        << "  <Author>" << Author << "</Author>" << eol 
        << "  <LastAuthor>" << LastAuthor << "</LastAuthor>" << eol 
        << "  <Created>2017-08-04T14:45:44Z</Created>" << eol 
        << "  <LastSaved>2017-08-04T15:23:17Z</LastSaved>" << eol 
        << "  <Version>16.00</Version>" << eol 
        << " </DocumentProperties>" << eol 
        << " <OfficeDocumentSettings xmlns=\"urn:schemas-microsoft-com:office:office\">" << eol 
        << "   <AllowPNG/>" << eol 
        << "  </OfficeDocumentSettings>" << eol 
        << "  <ExcelWorkbook xmlns=\"urn:schemas-microsoft-com:office:excel\">" << eol 
        << "   <WindowHeight>" << WindowHeight << "</WindowHeight>" << eol 
        << "   <WindowWidth>" << WindowWidth << "</WindowWidth>" << eol 
        << "   <WindowTopX>0</WindowTopX>" << eol 
        << "   <WindowTopY>0</WindowTopY>" << eol 
        << "   <ActiveSheet>" << ActiveSheet << "</ActiveSheet>" << eol 
        << "   <ProtectStructure>False</ProtectStructure>" << eol 
        << "   <ProtectWindows>False</ProtectWindows>" << eol 
        << "  </ExcelWorkbook>" << eol;

    return rv.str();
}

/// Output the styles tags.
const XMLSTR XmlWorkBook::stylesTags()
{
    std::stringstream rv; 
    XMLSTR eol = "\n";
    rv << " <Styles>" << eol;
    for (size_t i = 0; i < Styles.size(); i++)
    {
        rv << Styles.at(i).xml();
    }
    
    rv << " </Styles>" << eol;
    return rv.str();        
}

/// Worksheets
const XMLSTR XmlWorkBook::workSheets()
{
    std::stringstream rv; 
    for (size_t i = 0; i < Sheets.size(); i++)
    {
        rv << Sheets.at(i).xml();       
    }
    return rv.str();      
}

/// End tags.
const XMLSTR XmlWorkBook::endTags()
{
    return "</Workbook>";
}
