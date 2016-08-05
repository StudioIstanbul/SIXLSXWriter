//
//  SIXLSXWriter.h
//  exchangeExport
//
//  Created by Andreas Zöllner on 05/08/16.
//  Copyright © 2016 Studio Istanbul. All rights reserved.
//

#import <Foundation/Foundation.h>
#import "xlsxwriter.h"

typedef NS_ENUM(NSUInteger, SIXLSXChartType) {
    SIXLSXChartArea,
    SIXLSXChartAreaStacked,
    SIXLSXChartAreaStackedPercent,
    SIXLSXChartBar,
    SIXLSXChartBarStacked,
    SIXLSXChartBarStackedPercent,
    SIXLSXChartColumn,
    SIXLSXChartColumnStacked,
    SIXLSXChartColumnStackedPercent,
    SIXLSXChartDoughnut,
    SIXLSXChartLine,
    SIXLSXChartPie,
    SIXLSXChartScatter,
    SIXLSXChartScatterStraight,
    SIXLSXChartScatterStraightWithMarkers,
    SIXLSXChartScatterSmooth,
    SIXLSXChartScatterSmoothWithMarkers,
    SIXLSXChartRadar,
    SIXLSXChartRadarWithMarkers,
    SIXLSXChartRadarFilled
};

/**
 *  SIXLSXFormat class describes the format of a cell
 */

@interface SIXLSXFormat : NSObject

-(void)setFontName:(NSString*)fontName;
-(void)setFontSize:(int)fontSize;
-(void)setBold;
-(void)setItalic;
-(void)setUnderlined;
-(void)setStrikeout;
-(void)setFontColor:(NSColor*)color;
-(void)setBackgroundColor:(NSColor*)color;
@end

@interface SIXLSXChart : NSObject

@end

@interface SIXLSXWorksheet : NSObject

/**
 *  writes a string to a worksheet cell.
 *
 *  @param string         string to write
 *  @param cellIdentifier cell identifier in Excel style (ex. A1)
 *  @param cellFormat     format to use, nil for default.
 *
 *  @return YES if successful, NO if not.
 */

-(BOOL)writeString:(NSString*)string toCell:(NSString*)cellIdentifier withFormat:(SIXLSXFormat*)cellFormat;

-(BOOL)writeString:(NSString *)string toRow:(int) row andColumn:(int) column withFormat:(SIXLSXFormat *)cellFormat;

-(BOOL)setColumnWidth:(int)width forColumn:(int)startColumn toColumn:(int)endColumn andFormat:(SIXLSXFormat*)cellFormat;

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
 *  closes the current workbook.
 *
 *  @param error error if operation failes
 *
 *  @return YES if successfull, NO if error occured.
 */

-(BOOL)closeOrError:(NSError**)error;

/**
 *  sets document properties from a NSDictionary. Supported property keys are:
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
 *  all other keys will be added as custom property. Custom properties can be of type NSString, NSDate or NSNumber.
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