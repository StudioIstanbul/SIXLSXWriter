//
//  SIXLSXWriter.h
//  SIXLSXWriter
//
//  Created by Andreas Zöllner on 05/08/16.
//  Copyright © 2016 Studio Istanbul. All rights reserved.
//

#import <Cocoa/Cocoa.h>

//! Project version number for SIXLSXWriter.
FOUNDATION_EXPORT double SIXLSXWriterVersionNumber;

//! Project version string for SIXLSXWriter.
FOUNDATION_EXPORT const unsigned char SIXLSXWriterVersionString[];

// In this header, you should import all the public headers of your framework using statements like #import <SIXLSXWriter/PublicHeader.h>

//#import "xlsxwriter.h"

/**
 *  Defines options for a row or column. Only hidden property is supported at the moment.
 */
@interface SIXLSXRowColOptions : NSObject

/**
 *  The "hidden" option is used to hide a column. This can be used, for example, to hide intermediary steps in a complicated calculation.
 */
@property (assign) BOOL hidden;

/**
 *  This property is not supported yet.
 */
@property (assign) BOOL collapsed;

/**
 *  This property is not supported yet.
 */
@property (assign) BOOL level;

/**
 *  Returns an initialized options object. The "hidden" option is used to hide a column. This can be used, for example, to hide intermediary steps in a complicated calculation.
 *
 *  @param hidden    The "hidden" option is used to hide a column. This can be used, for example, to hide intermediary steps in a complicated calculation.
 *  @param collapsed This property is not supported yet.
 *  @param level     This property is not supported yet.
 *
 *  @return an initialized options object
 */
+(SIXLSXRowColOptions*)rowColOptionsWithHidden:(BOOL)hidden collapsed:(BOOL)collapsed andLevel:(BOOL)level;

@end

/**
 *  SIXLSXFormat class describes the format of a cell
 */

@interface SIXLSXFormat : NSObject

/**
 *  Sets the font of a format.
 *
 *  @param fontName name of the font to use.
 */
-(void)setFontName:(NSString*)fontName;

/**
 *  Sets the size of a font for a format.
 *
 *  @param fontSize font size to use
 */
-(void)setFontSize:(int)fontSize;

/**
 *  Sets the font to bold.
 */
-(void)setBold;

/**
 *  Sets the font to italic.
 */
-(void)setItalic;

/**
 *  Sets the text to underlined.
 */
-(void)setUnderlined;

/**
 *  Sets the text to strikeout.
 */
-(void)setStrikeout;

/**
 *  Sets the color of the text.
 *
 *  @param color color to use for text
 */
-(void)setFontColor:(NSColor*)color;

/**
 *  Sets the background color of the cell.
 *
 *  @param color color to use for background.
 */
-(void)setBackgroundColor:(NSColor*)color;

/** Sets the format for number cell contents.

This method is used to define the numerical format of a number in Excel. It controls whether a number is displayed as an integer, a floating point number, a date, a currency value or some other user defined format.
 
The numerical format of a cell can be specified by using a format string:

Format strings can control any aspect of number formatting allowed by Excel:

Examples for valid number formats:

        @"0.000
        @"#,##0"
        @"#,##0.00
        @"0.00"
        @"mm/dd/yy"
        @"mmm d yyyy
        @"d mmmm yyyy"
        @"dd/mm/yyyy hh:mm AM/PM"
        @"0 \"dollar and\" .00 \"cents\""
        @"[Green]General;[Red]-General;General"
        @"00000"

The number system used for dates is described in Working with Dates and Times.
 
For more information on number formats in Excel refer to the Microsoft documentation on cell formats.

@param numberFormat string representing number format to use.
*/
-(void)setNumberFormat:(NSString*)numberFormat;

@end

/**
 *  Defines a chart object.
 */

@interface SIXLSXChart : NSObject

/**
 *  Supported type of charts.
 */
typedef NS_ENUM(NSUInteger, SIXLSXChartType) {
    /**
     *  Area chart.
     */
    SIXLSXChartArea,
    /**
     *  Area chart - stacked.
     */
    SIXLSXChartAreaStacked,
    /**
     *  Area chart - percentage stacked.
     */
    SIXLSXChartAreaStackedPercent,
    /**
     *  Bar chart.
     */
    SIXLSXChartBar,
    /**
     *  Bar chart - stacked.
     */
    SIXLSXChartBarStacked,
    /**
     *  Bar chart - percentage stacked.
     */
    SIXLSXChartBarStackedPercent,
    /**
     *  Column chart.
     */
    SIXLSXChartColumn,
    /**
     *  Column chart - stacked.
     */
    SIXLSXChartColumnStacked,
    /**
     * Column chart - percentage stacked.
     */
    SIXLSXChartColumnStackedPercent,
    /**
     *  Doughnut chart.
     */
    SIXLSXChartDoughnut,
    /**
     *  Line chart.
     */
    SIXLSXChartLine,
    /**
     *  Pie chart.
     */
    SIXLSXChartPie,
    /**
     *  Scatter chart.
     */
    SIXLSXChartScatter,
    /**
     *  Scatter chart - straight.
     */
    SIXLSXChartScatterStraight,
    /**
     *  Scatter chart - straight with markers.
     */
    SIXLSXChartScatterStraightWithMarkers,
    /**
     *  Scatter chart - smooth.
     */
    SIXLSXChartScatterSmooth,
    /**
     *  Scatter chart - smooth with markers.
     */
    SIXLSXChartScatterSmoothWithMarkers,
    /**
     *  Radar chart.
     */
    SIXLSXChartRadar,
    /**
     *  Radar chart - with markers.
     */
    SIXLSXChartRadarWithMarkers,
    /**
     *  Radar chart - filled.
     */
    SIXLSXChartRadarFilled
};


@end

/**
 *  Defines a worksheet contained in the files workbook.
 */

@interface SIXLSXWorksheet : NSObject

/**
 *  writes a string to a worksheet cell.
 *
 *  @param string         string to write
 *  @param cellIdentifier cell identifier in Excel style (ex. @"A1")
 *  @param cellFormat     format to use, nil for default.
 *
 *  @return YES if successful, NO if not.
 */

-(BOOL)writeString:(NSString*)string toCell:(NSString*)cellIdentifier withFormat:(SIXLSXFormat*)cellFormat;

/**
 *  writes a string to a worksheet cell.
 *
 *  @param string     writes a string to a worksheet cell.
 *  @param row        row of cell
 *  @param column     column of cell
 *  @param cellFormat format to use, nil for default.
 *
 *  @return YES if successful, NO if not.
 */

-(BOOL)writeString:(NSString *)string toRow:(int) row andColumn:(int) column withFormat:(SIXLSXFormat *)cellFormat;

/**
 *  Sets width and format of a column.
 *
 *  This method can be used to change the default properties of a single column or a range of columns. If this is applied to a single column the value of first_col and last_col should be the same. The width parameter sets the column width in the same units used by Excel which is: the number of characters in the default font. The default width is 8.43 in the default font of Calibri 11. The actual relationship between a string width and a column width in Excel is complex. See the following explanation of column widths from the Microsoft support documentation for more details.
 *  There is no way to specify "AutoFit" for a column in the Excel file format. This feature is only available at runtime from within Excel. It is possible to simulate "AutoFit" in your application by tracking the maximum width of the data in the column as your write it and then adjusting the column width at the end.
 *
 *  As usual the format parameter is optional. If you wish to set the format without changing the width you can pass a default column width of 8.43.
 *
 *  The format parameter will be applied to any cells in the column that don't have a format. For example:
 *
 *  As in Excel a row format takes precedence over a default column format.
 *
 *  @param width       width of the column in number of characters in the default font (Calibri 11pt)
 *  @param startColumn index of column to start
 *  @param endColumn   index of column to end
 *  @param cellFormat  format to apply or nil
 *
 *  @return YES if successful, NO if not.
 */

/**
 *  Writes a number to a cell on your worksheet.
 *
 *  The native data type for all numbers in Excel is a IEEE-754 64-bit double-precision floating point, which is also the default type used.
 *
 *  The cellFormat parameter is used to apply formatting to the cell. This parameter can be nil to indicate no formatting or it can be a SIXLSXFormat object.
 *
 *  @param number         numeric value to write
 *  @param cellIdentifier cell identifier in Excel style (ex. @"A1")
 *  @param cellFormat     format to use, nil for default.
 *
 *  @return YES if successful, NO if not.
 */
-(BOOL)writeNumber:(NSNumber*)number toCell:(NSString*)cellIdentifier withFormat:(SIXLSXFormat*)cellFormat;

/**
 *  Writes a number to a cell on your worksheet.
 *
 *  The native data type for all numbers in Excel is a IEEE-754 64-bit double-precision floating point, which is also the default type used.
 *
 *  The cellFormat parameter is used to apply formatting to the cell. This parameter can be nil to indicate no formatting or it can be a SIXLSXFormat object.
 *
 *  @param number     numeric value to write
 *  @param row        row index of the cell
 *  @param column     column index of the cell
 *  @param cellFormat format to apply or nil
 *
 *  @return YES if successfull, NO if not.
 */
-(BOOL)writeNumber:(NSNumber *)number toRow:(int)row andColumn:(int)column withFormat:(SIXLSXFormat *)cellFormat;

/**
 *  Writes a date to a cell on your worksheet.
 *
 *  The format parameter should be used to apply formatting to the cell using a Format object as shown above. Without a date format the datetime will appear as a number only.
 *
 *  @param date             date value to write
 *  @param cellIdentifier   cell identifier in Excel style (ex. @"A1")
 *  @param cellFormat       format to use, nil for default.
 *
 *  @return YES if successful, NO if not.
 */
-(BOOL)writeDate:(NSDate *)date toCell:(NSString *)cellIdentifier withFormat:(SIXLSXFormat *)cellFormat;

/**
 *  Writes a date to a cell on your worksheet.
 *
 *  The format parameter should be used to apply formatting to the cell using a Format object as shown above. Without a date format the datetime will appear as a number only.
 *
 *  @param date             date value to write
 *  @param row              index of cell row
 *  @param column           index of column
 *  @param cellFormat       format to use, nil for default.
 *
 *  @return YES if successful, NO if not.
 */
-(BOOL)writeDate:(NSDate *)date toRow:(int)row andColumn:(int)column withFormat:(SIXLSXFormat *)cellFormat;

/**
 *  Writes a formula to a cell on your worksheet.
 *
 *    @"=B3 + 6"
 *    @"=SIN(PI()/4)"
 *    @"=SUM(A1:A2)"
 *    @"=IF(A3>1,\"Yes\", \"No\")"
 *    @"=AVERAGE(1, 2, 3, 4)"
 *    @"=DATEVALUE(\"1-Jan-2013\")"
 *
 *  Libxlsxwriter doesn't calculate the value of a formula and instead stores a default value of 0. The correct formula result is displayed in Excel, as shown in the example above, since it recalculates the formulas when it loads the file. For cases where this is an issue see the worksheet_write_formula_num() function and the discussion in that section.
 *
 *  Formulas must be written with the US style separator/range operator which is a comma (not semi-colon).
 *
 *  @param formula        formula to use
 *  @param cellIdentifier cell identifier in Excel style (ex. @"A1")
 *  @param cellFormat     format to use, nil for default.
 *
 *  @return YES if successful, NO if not.
 */
-(BOOL)writeFormula:(NSString*)formula toCell:(NSString*)cellIdentifier withFormat:(SIXLSXFormat*)cellFormat;

/**
 *  Writes a formula to a cell on your worksheet.
 *
 *    @"=B3 + 6"
 *    @"=SIN(PI()/4)"
 *    @"=SUM(A1:A2)"
 *    @"=IF(A3>1,\"Yes\", \"No\")"
 *    @"=AVERAGE(1, 2, 3, 4)"
 *    @"=DATEVALUE(\"1-Jan-2013\")"
 *
 *  Libxlsxwriter doesn't calculate the value of a formula and instead stores a default value of 0. The correct formula result is displayed in Excel, as shown in the example above, since it recalculates the formulas when it loads the file. For cases where this is an issue see the -writeFormula:withResult:toRow:andColumn:withFormat method and the discussion in that section.
 *
 *  Formulas must be written with the US style separator/range operator which is a comma (not semi-colon).
 *
 *  @param formula    formula to use
 *  @param row        index of cell row
 *  @param column     index of cell column
 *  @param cellFormat format to use, nil for default.
 *
 *  @return YES if successful, NO if not.
 *  @see -writeFormula:toRow:andColumn:withFormat:
 */

-(BOOL)writeFormula:(NSString *)formula toRow:(int)row andColumn:(int)column withFormat:(SIXLSXFormat *)cellFormat;

/**
 *  This method writes a formula or Excel function to the cell specified by row and column with a user defined result.
 *
 *  Libxlsxwriter doesn't calculate the value of a formula and instead stores the value 0 as the formula result. It then sets a global flag in the XLSX file to say that all formulas and functions should be recalculated when the file is opened.
 *
 *  This is the method recommended in the Excel documentation and in general it works fine with spreadsheet applications.
 *
 *  However, applications that don't have a facility to calculate formulas, such as Excel Viewer, QuickView or some mobile applications will only display the 0 results.
 *
 *  If required, this method can be used to specify a formula and its result.
 *
 *  This function is rarely required and is only provided for compatibility with some third party applications. For most applications the -writeFormula:toCell:withFormat: method is the recommended way of writing formulas.
 *
 *  @param formula    formula to use
 *  @param result     calculation result to use
 *  @param row        index of cell row
 *  @param column     index of cell column
 *  @param cellFormat format to use, nil for default.
 *
 *  @return YES if successful, NO if not.
 *  @see -writeFormula:toCell:withFormat:
 */
-(BOOL)writeFormula:(NSString*)formula withResult:(NSNumber*)result toRow:(int)row andColumn:(int)column withFormat:(SIXLSXFormat*)cellFormat;

/**
 *  This method writes a formula or Excel function to the cell specified by row and column with a user defined result.
 *
 *  Libxlsxwriter doesn't calculate the value of a formula and instead stores the value 0 as the formula result. It then sets a global flag in the XLSX file to say that all formulas and functions should be recalculated when the file is opened.
 *
 *  This is the method recommended in the Excel documentation and in general it works fine with spreadsheet applications.
 *
 *  However, applications that don't have a facility to calculate formulas, such as Excel Viewer, QuickView, or some mobile applications will only display the 0 results.
 *
 *  If required, this method can be used to specify a formula and its result.
 *
 *  This function is rarely required and is only provided for compatibility with some third party applications. For most applications the -writeFormula:toCell:withFormat: method is the recommended way of writing formulas.
 *
 *  @param formula          formula to use
 *  @param result           calculation result to use
 *  @param cellIdentifier   cell identifier in Excel style (ex. @"A1")
 *  @param cellFormat format to use, nil for default.
 *
 *  @return YES if successful, NO if not.
 *  @see -writeFormula:toCell:withFormat:
 */
-(BOOL)writeFormula:(NSString*)formula withResult:(NSNumber*)result toCell:(NSString*)cellIdentifier withFormat:(SIXLSXFormat*)cellFormat;

/**
 *  Sets width and format of a column.
 *
 *  This method can be used to change the default properties of a single column or a range of columns. If this is applied to a single column the value of first_col and last_col should be the same. The width parameter sets the column width in the same units used by Excel which is: the number of characters in the default font. The default width is 8.43 in the default font of Calibri 11. The actual relationship between a string width and a column width in Excel is complex. See the following explanation of column widths from the Microsoft support documentation for more details.
 *  There is no way to specify "AutoFit" for a column in the Excel file format. This feature is only available at runtime from within Excel. It is possible to simulate "AutoFit" in your application by tracking the maximum width of the data in the column as your write it and then adjusting the column width at the end.
 *
 *  As usual the format parameter is optional. If you wish to set the format without changing the width you can pass a default column width of 8.43.
 *
 *  The format parameter will be applied to any cells in the column that don't have a format. For example:
 *
 *  As in Excel a row format takes precedence over a default column format.
 *
 *  @param width       width of the column in number of characters in the default font (Calibri 11pt)
 *  @param startColumn index of the first column
 *  @param endColumn   index of the last column
 *  @param cellFormat  ormat to apply or nil
 *  @param options     options to apply
 *
 *  @return YES if successful, NO if not.
 */
-(BOOL)setColumnWidth:(int)width forColumn:(int)startColumn toColumn:(int)endColumn andFormat:(SIXLSXFormat *)cellFormat withOptions:(SIXLSXRowColOptions*)options;

/**
 *  Sets width and format of a column.
 *
 *  This method can be used to change the default properties of a single column or a range of columns. If this is applied to a single column the value of first_col and last_col should be the same. The width parameter sets the column width in the same units used by Excel which is: the number of characters in the default font. The default width is 8.43 in the default font of Calibri 11. The actual relationship between a string width and a column width in Excel is complex. See the following explanation of column widths from the Microsoft support documentation for more details.
 *  There is no way to specify "AutoFit" for a column in the Excel file format. This feature is only available at runtime from within Excel. It is possible to simulate "AutoFit" in your application by tracking the maximum width of the data in the column as your write it and then adjusting the column width at the end.
 *
 *  As usual the format parameter is optional. If you wish to set the format without changing the width you can pass a default column width of 8.43.
 *
 *  The format parameter will be applied to any cells in the column that don't have a format. For example:
 *
 *  As in Excel a row format takes precedence over a default column format.
 *
 *  @param width       width of the column in number of characters in the default font (Calibri 11pt)
 *  @param startColumn index of the first column
 *  @param endColumn   index of the last column
 *  @param cellFormat  format to apply or nil
 *
 *  @return YES if successful, NO if not.
 */

-(BOOL)setColumnWidth:(int)width forColumn:(int)startColumn toColumn:(int)endColumn andFormat:(SIXLSXFormat*)cellFormat;

/**
 *  Sets width and format of a column.
 *
 *  This method can be used to change the default properties of a single column or a range of columns. If this is applied to a single column the value of first_col and last_col should be the same. The width parameter sets the column width in the same units used by Excel which is: the number of characters in the default font. The default width is 8.43 in the default font of Calibri 11. The actual relationship between a string width and a column width in Excel is complex. See the following explanation of column widths from the Microsoft support documentation for more details.
 *  There is no way to specify "AutoFit" for a column in the Excel file format. This feature is only available at runtime from within Excel. It is possible to simulate "AutoFit" in your application by tracking the maximum width of the data in the column as your write it and then adjusting the column width at the end.
 *
 *  As usual the format parameter is optional. If you wish to set the format without changing the width you can pass a default column width of 8.43.
 *
 *  The format parameter will be applied to any cells in the column that don't have a format. For example:
 *
 *  As in Excel a row format takes precedence over a default column format.
 *
 *  @param width       width of the column in number of characters in the default font (Calibri 11pt)
 *  @param columnRange range of columns to be affected in Excel style syntax (ex. "A:E")
 *  @param cellFormat  format to apply or nil
 *
 *  @return YES if successful, NO if not.
 */

-(BOOL)setColumnWidth:(int)width forColumnRange:(NSString*)columnRange andFormat:(SIXLSXFormat *)cellFormat;

@end

/**
 *  SIXLSXWorkbookOptions describes the options used for handling workbook files.
 */

@interface SIXLSXWorkbookOptions : NSObject

/**
 *  Reduces the amount of data stored in memory so that large files can be written efficiently.
 
 Note
 In this mode a row of data is written and then discarded when a cell in a new row is added. Therefore, once this option is active, data should be written in sequential row order. For this reason the worksheet_merge_range() doesn't work in this mode.
 */

@property (assign) BOOL constantMemory;

/**
 *  NSURL of directory for storing temporary files. libxlsxwriter stores workbook data in temporary files prior to assembling the final XLSX file. The temporary files are created in the system's temp directory. If the default temporary directory isn't accessible to your application, or doesn't contain enough space, you can specify an alternative location using the tempdir option.
 */

@property (strong) NSURL* tmpdir;

@end

/**
 *  SIXLSXWorkbook is the workbook class and therefore the main class for creating a new workbook file.
 */

@interface SIXLSXWorkbook : NSObject

/**
 *  create a new workbook
 *
 *  @param fileURL the NSURL where to store this workbook, must not be nil and a valid URL
 *
 *  @return the newly created workbook
 */

+(SIXLSXWorkbook*)newWorkbookAtURL:(NSURL*)fileURL;

/**
 *  creates a new workbook with specific options.
 *
 *  @param options the options to use for this workbook or nil.
 *  @param fileURL the NSURL where to store this workbook, must not be nil and a valid URL
 *
 *  @see SIXLSXWorkbookOptions
 *
 *  @return the newly created workbook
 */

+(SIXLSXWorkbook*)newWorkbookWithOptions:(SIXLSXWorkbookOptions*)options atURL:(NSURL*)fileURL;

/**
 *  initializes a new workbook
 *
 *  @param fileUrl fileURL the NSURL where to store this workbook, must not be nil and a valid URL
 *
 *  @return the initialized workbook
 */

-(SIXLSXWorkbook*)initWithURL:(NSURL*)fileUrl;

/**
 *  initializes a new workbook with specific options.
 *
 *  @param fileUrl the NSURL where to store this workbook, must not be nil and a valid URL
 *  @param options the options to use for this workbook or nil.
 *
 *  @return the initialized workbook
 */

-(SIXLSXWorkbook*)initWithURL:(NSURL *)fileUrl andOptions:(SIXLSXWorkbookOptions*)options;

/**
 *  adds a new worksheet to this workbook. The worksheet name must be a valid Excel worksheet name, i.e. it must be less than 32 character and it cannot contain any of the characters:
 *
 *  / \ [ ] : * ?
 *  In addition, you cannot use the same, case insensitive, sheetname for more than one worksheet.
 *
 *  @param worksheetTitle the title for the new worksheet or nil.
 *
 *  @return the newly added and initialized worksheet.
 */

-(SIXLSXWorksheet*)addWorksheetWithTitle:(NSString*)worksheetTitle;

/**
 *  creates a new format to be used for formatting cells.
 *
 *  @return the newly added format.
 */

-(SIXLSXFormat*)addFormat;

/**
 *  adds a new chart of the desired type to the existing workbook.
 *
 *  @param chartType type of chart
 *
 *  @return the newly added chart
 */

-(SIXLSXChart*)addChartOfType:(SIXLSXChartType)chartType;


/**
 *  Closes the current workbook.
 *
 *  @param error error if operation failes
 *
 *  @return YES if successfull, NO if error occured.
 */

-(BOOL)closeOrError:(NSError**)error;

/**
 *  sets document properties from a NSDictionary.
 *
 *  Supported property keys are:
 *
 *  - title
 *  - subject
 *  - author
 *  - manager
 *  - company
 *  - category
 *  - keywords
 *  - comments
 *  - hyperlink_base
 *
 *  All other keys will be added as custom property. Custom properties can be of type NSString, NSDate or NSNumber.
 *
 *  @param propertiesDictionary the dictionary of properties to use
 *
 *  @return YES if successfull, NO if error occured.
 */

-(BOOL)setDocumentPropertiesFromDictionary:(NSDictionary*)propertiesDictionary;

/**
 *  all worksheets in this workbook
 *
 *  @return array of all worksheets
 */

-(NSArray*)worksheets;

@end