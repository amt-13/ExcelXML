/**
 * Copyright 2017, Philip R. Braica (HoshiKata@aol.com, VeryMadSci@gmail.com)
 *   Distributed under the The Code Project Open License (CPOL)
 *   http://www.codeproject.com/info/cpol10.aspx
 *
 * @brief This software unit can generate .xml spread sheet files  
 *        consumable by Office and OpenOffice. It is usable example 
 *        code, with some basic features and doxygen.
 *
 * Note, though the author is aware of how nice C++ 11 semantics 
 * (shared_ptr) would make the code, it was purposefully not done 
 * because not all projects allow for C++ 11. This is meant to be 
 * minimal, reasonable to use (speed, memory) with few dependencies, 
 * and easy to port for embedded, desktop, and to C# / python. Hence, 
 * structs not classes, no accessors, no weak/shared pointers, no 
 * boost, very plain C++.
 *
 * It purely captures the confusing parts of generating the  file 
 * in a way that is reference, not performance, cool linguistically, 
 * nor pretty.
 */

#ifndef XMLWORKBOOK_H_      ///< Check for inclusion guard.
#define XMLWORKBOOK_H_      ///< Define inclusion guard.

// Includes.
#include <string>
#include <vector>

/// Define XMLSTR as std::string, save some typing and allow
/// some easy porting, say to QString or your own companies 
/// string wrapper.
typedef std::string XMLSTR;

/// A Cell style.
class XmlStyle
{
public:
    // Data.
    XMLSTR ID;             ///< "Default" or "s1" etc. do not enclose in quotes.
    XMLSTR Name;           ///< "Normal" or leave blank.
    XMLSTR Alignment;      ///< ss::Vertical=\"Bottom\", etc.
    XMLSTR Font;           ///< ss::FontName=\"Calibri\" x:Family=\"Swiss\" ss:Size=\"11\" ss:Color="#000000"
    XMLSTR Borders;        ///< see setBorders
    XMLSTR Interior;       ///< ss:Color=\"#FFFF00\" ss:Pattern=\"Solid\"
    XMLSTR NumberFormat;
    XMLSTR Protection;    

    enum HorizontalAlignment
    {
        AlignHLeft,
        AlignHCenter,
        AlignHRight
    };

    enum VerticalAlignment
    {
        AlignVTop,
        AlignVCenter,
        AlignVBottom
    };

    // Primary functions.
    XmlStyle(const XMLSTR & name = "Default"); ///< Construct as default style.
    virtual ~XmlStyle();                            ///< Destructor.
    const XMLSTR xml();                             ///< Output as XML.
    const XMLSTR xmlwithoutIDName();      /// < Output XML for Comparision

    // API functions to make use easier.

    /// Set the font string.
    void setFont
        (
            const XMLSTR & fontName = "Calibri",    ///< Font name.
            const XMLSTR & fontFamily = "Swiss",    ///< Family name.
            float size = 11,                        ///< Size in points.
            long color = 0x000000                   ///< RGB 24 bit color.
        );

    /// Append bold.
    void appendFontBold()            { Font += " ss:Bold=\"1\""; }          

    /// Append Italic.
    void appendFontItalic()          { Font += " ss:Italic=\"1\""; }

    // Append single underline.
    void appendFontSingleUnderline() { Font += " ss:Underline=\"Single\""; }

    // Append double underline.
    void appendFontDoubleUnderline() { Font += " ss:Underline=\"Double\""; }

    /// Set the interior string.
    void setInterior
        (
            long color = 0x000000,              ///< 24 bit RBG color.
            const XMLSTR & pattern = "Solid"    ///< Fill pattern.
        );
    
    ///< Create simple borders with the give, weights, 0=off.
    void setBorders
        (
            size_t topWeight = 0,       ///< Top weight or off.
            size_t bottomWeight = 0,    ///< Bottom weight or off.
            size_t leftWeight = 0,      ///< Left weight or off.
            size_t rightWeight = 0      ///< Right weight or off.
        );

    ///< Create Alignment.
    void setAlignment
        (
            HorizontalAlignment hAlign = AlignHLeft,       ///< AlignHLeft.
            VerticalAlignment vAlign = AlignVBottom,    ///< AlignVTop.
            bool mWrapText = false,                      ///< Text Wrap.
            int mTextRotate = 0

        );

    bool operator== (XmlStyle& s1);
    bool operator!= (XmlStyle& s1);
};

/// cell in row.
class XMLCell
{
public:
    // Data.
    XMLSTR Style;      ///< Style ID  if not Normal.
    XMLSTR Type;       ///< String or Number
    XMLSTR Value;      ///< Contents.
    XMLSTR Formula;    ///< SUM(RC[-1]:R[3]C[-1])
    size_t Index;      ///< Auto increment is zero, else the column, A=1, B=2, etc.
    size_t mergeCols = 0; ///< ss:MergeAcross="3"
    size_t mergeRows = 0; ///< ss:MergeAcross="3"

    // Primary functions.
    XMLCell();              ///< Constructor.
    virtual ~XMLCell();     ///< Destructor.
    const XMLSTR xml();     ///< Output as XML.

    // API functions for ease of use.

    /// Assign style.
    virtual void setStyle
        (
            const XmlStyle & style          ///< Style to use.
        )
    {
        Style = style.ID;
    }

    /// Assign style.
    virtual void setStyle
        (
            const XmlStyle * const style    ///< Style to use.
        )
    {
        Style = style->ID;
    }

};

/// Row in table.
class XMLRow
{
public:
    // Data.
    std::vector<XMLCell>    Cells;  ///< Cells

    // Primary functions.
    XMLRow();                       ///< Constructor.
    virtual ~XMLRow();              ///< Destructor.
    const XMLSTR xml();             ///< Output as XML.
};

/// Table in worksheet.
class XMLTable
{
public:
    // Data.
    float RowHeight;                ///< 15 for default.
    std::vector<XMLRow> Rows;       ///< The rows.

    // Primary functions.
    XMLTable();                     ///< Constructor.
    virtual ~XMLTable();            ///< Destructor.
    const XMLSTR xml();             ///< Output as XML.

    /// Shift a row up or down.
    void insertRow
        (
            size_t whichRow,        ///< Which row.
            XMLRow * row = NULL     ///< Row or null to insert a blank row.
        );
};

/// A worksheet structure.
class XMLWorkSheet
{
public:
    // Data.
    XMLSTR                  Name;   ///< Sheet name.
    std::vector<XMLTable>   Tables; ///< Tables

    // Primary functions.
    XMLWorkSheet();                 ///< Constructor.
    virtual ~XMLWorkSheet();        ///< Destructor.
    const XMLSTR xml();             ///< Output as XML.

    // Utility functions.

    /// Create a worksheet with a table and cols/rows.
    void addTable
        (
            XMLSTR * colNames,      ///< Column names of size cols or null.
            XMLSTR * colNameStyle,  ///< Style of column names.
            XMLSTR * dataStyle,     ///< Data style name.
            size_t rows,            ///< Number of rows.
            size_t cols,            ///< Number of columns.
            double * pData,         ///< rows*cols data. 
            bool asRows = true,     ///< True means data is [r0,c0] [r0,c1]... [r0,cn],[r1,c0] ... else by column.
            size_t startRow = 0,    ///< Start Row offset.
            size_t startCol = 0     ///< Start Column offset.
        );
};

/// A workbook.
class XmlWorkBook
{
public:
    // Data.
    XMLSTR                      Author;         ///< Original author.
    XMLSTR                      LastAuthor;     ///< last author.
    size_t                      ActiveSheet;    ///< Which sheet is on top.
    size_t                      WindowWidth;    ///< Default window width on open.
    size_t                      WindowHeight;   ///< Default window height on open.
    std::vector<XMLWorkSheet>   Sheets;         ///< List of worksheets.
    std::vector<XmlStyle>       Styles;         ///< List of styles.

    // Primary functions.
    XmlWorkBook();              ///< Constructor.
    virtual ~XmlWorkBook();     ///< Destructor.
    const XMLSTR xml();         ///< Header tags, styles, worksheets, endtags.

    // Utility if output needs customization.
    const XMLSTR headerTags();  ///< Output the xml, workbook, document properties, etc. tags.    
    const XMLSTR stylesTags();  ///< Output the styles tags.
    const XMLSTR workSheets();  ///< Worksheets   
    const XMLSTR endTags();     ///< End tags.
};

#endif XMLWORKBOOK_H_
