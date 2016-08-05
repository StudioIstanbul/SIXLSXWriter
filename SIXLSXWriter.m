//
//  SIXLSXWriter.m
//  exchangeExport
//
//  Created by Andreas Zöllner on 05/08/16.
//  Copyright © 2016 Studio Istanbul. All rights reserved.
//

#import "SIXLSXWriter.h"

@interface SIXLSXFormat () {
    lxw_format* _lxwFormat;
}
+(SIXLSXFormat*)_formatFromAdd:(lxw_format*)format;
-(lxw_format*)_lxwFormat;
@end

@implementation SIXLSXFormat

-(SIXLSXFormat*)initWithFormat:(lxw_format*)format {
    self = [super init];
    if (self) {
        _lxwFormat = format;
    }
    return self;
}

+(SIXLSXFormat*)_formatFromAdd:(lxw_format *)format {
    return [[SIXLSXFormat alloc] initWithFormat:format];
}

-(lxw_format*)_lxwFormat {
    return _lxwFormat;
}

-(void)setFontName:(NSString *)fontName {
    format_set_font_name(_lxwFormat, [fontName cStringUsingEncoding:NSUTF8StringEncoding]);
}

-(void)setFontSize:(int)fontSize {
    format_set_font_size(_lxwFormat, fontSize);
}

-(void)setBold {
    format_set_bold(_lxwFormat);
}

-(void)setItalic {
    format_set_italic(_lxwFormat);
}

-(void)setUnderlined {
    format_set_underline(_lxwFormat, LXW_UNDERLINE_SINGLE);
}

-(void)setStrikeout {
    format_set_font_strikeout(_lxwFormat);
}

-(void)setFontColor:(NSColor *)color {
    uint32 col = 0;
    if ([color.colorSpaceName isEqualToString:@"NSCalibratedWhiteColorSpace"]) col = (color.whiteComponent * 255 * 256 * 256) + (color.whiteComponent * 255 * 256) + color.whiteComponent * 255;
    if ([color.colorSpaceName isEqualToString:@"NSCalibratedRGBColorSpace"]) col = (color.redComponent * 255 * 256 * 256) + (color.greenComponent * 255 * 256) + color.blueComponent * 255;
    format_set_font_color(_lxwFormat, col);
}

-(void)setBackgroundColor:(NSColor *)color {
    uint32 col = 0;
    if ([color.colorSpaceName isEqualToString:@"NSCalibratedWhiteColorSpace"]) col = (color.whiteComponent * 255 * 256 * 256) + (color.whiteComponent * 255 * 256) + color.whiteComponent * 255;
    if ([color.colorSpaceName isEqualToString:@"NSCalibratedRGBColorSpace"]) col = (color.redComponent * 255 * 256 * 256) + (color.greenComponent * 255 * 256) + color.blueComponent * 255;
    format_set_bg_color(_lxwFormat, col);
}

@end

@interface SIXLSXChart () {
    lxw_chart* _lxwChart;
}
+(SIXLSXChart*)_chartFromAdd:(lxw_chart*)chart;

@end

@implementation SIXLSXChart

-(SIXLSXChart*)initWithChart:(lxw_chart*)chart {
    self = [super init];
    if (self) {
        _lxwChart = chart;
    }
    return self;
}

+(SIXLSXChart*)_chartFromAdd:(lxw_chart *)chart {
    return [[SIXLSXChart alloc] initWithChart:chart];
}

@end

@interface SIXLSXWorksheet () {
    lxw_worksheet* _lxwWorksheet;
}
+(SIXLSXWorksheet*)_worksheetFromAdd:(lxw_worksheet*)worksheet;

@end

@implementation SIXLSXWorksheet

-(id)initWithLwxWorksheet:(lxw_worksheet*)worksheet {
    self = [super init];
    if (self) {
        _lxwWorksheet = worksheet;
    }
    return self;
}

+(SIXLSXWorksheet*)_worksheetFromAdd:(lxw_worksheet *)worksheet {
    SIXLSXWorksheet* ws = [[SIXLSXWorksheet alloc] initWithLwxWorksheet:worksheet];
    return ws;
}

-(BOOL)writeString:(NSString *)string toCell:(NSString *)cellIdentifier withFormat:(SIXLSXFormat *)cellFormat {
    int row = lxw_name_to_row([cellIdentifier cStringUsingEncoding:NSASCIIStringEncoding]);
    int cell = lxw_name_to_col([cellIdentifier cStringUsingEncoding:NSASCIIStringEncoding]);
    return [self writeString:string toRow:row andColumn:cell withFormat:cellFormat];
}

-(BOOL)writeString:(NSString *)string toRow:(int)row andColumn:(int)column withFormat:(SIXLSXFormat *)cellFormat {
    lxw_error err =  worksheet_write_string(_lxwWorksheet, row, column, [string cStringUsingEncoding:NSUTF8StringEncoding], (cellFormat ? [cellFormat _lxwFormat] : NULL));
    if (err) return NO;
    return YES;
}

-(BOOL)setColumnWidth:(int)width forColumn:(int)startColumn toColumn:(int)endColumn andFormat:(SIXLSXFormat *)cellFormat {
    lxw_error err = worksheet_set_column(_lxwWorksheet, startColumn, endColumn, width, cellFormat._lxwFormat);
    if (err) return NO;
    return YES;
}

@end

@interface SIXLSXWorkbookOptions ()

-(lxw_workbook_options*)_lxwWorkbookOptions;

@end

@implementation SIXLSXWorkbookOptions

-(lxw_workbook_options*)_lxwWorkbookOptions {
    NSString* tmpDirStr = [self.tmpdir path];
    char sz[tmpDirStr.length*4];
    sprintf(sz, "%s", [tmpDirStr cStringUsingEncoding:NSUTF8StringEncoding]);
    lxw_workbook_options options = {.constant_memory = 0,
        .tmpdir = sz};
    if (self.constantMemory) options.constant_memory = 1;
    return &options;
}

@end

@interface SIXLSXWorkbook () {
    lxw_workbook* _lxwWorkbook;
    NSMutableArray* _worksheets;
}

@end


@implementation SIXLSXWorkbook

-(SIXLSXWorkbook*)initWithURL:(NSURL *)fileUrl andOptions:(SIXLSXWorkbookOptions *)options {
    self = [super init];
    if (self) {
        _worksheets = [[NSMutableArray alloc] init];
        if (options) {
            _lxwWorkbook = workbook_new_opt([[fileUrl path] cStringUsingEncoding:NSUTF8StringEncoding], [options _lxwWorkbookOptions]);
        } else {
            _lxwWorkbook = workbook_new([[fileUrl path] cStringUsingEncoding:NSUTF8StringEncoding]);
        }
    }
    return self;
}

-(SIXLSXWorkbook*)initWithURL:(NSURL *)fileUrl {
    return [self initWithURL:fileUrl andOptions:nil];
}

+(SIXLSXWorkbook*)newWorkbookWithOptions:(SIXLSXWorkbookOptions *)options atURL:(NSURL *)fileURL {
    SIXLSXWorkbook* wb = [[SIXLSXWorkbook alloc] initWithURL:fileURL andOptions:options];
    return wb;
}

+(SIXLSXWorkbook*)newWorkbookAtURL:(NSURL *)fileURL {
    return [SIXLSXWorkbook newWorkbookWithOptions:nil atURL:fileURL];
}

-(SIXLSXWorksheet*)addWorksheetWithTitle:(NSString *)worksheetTitle {
    lxw_worksheet* lxwWs;
    if (worksheetTitle && worksheetTitle.length <= 32) lxwWs = workbook_add_worksheet(_lxwWorkbook, [worksheetTitle cStringUsingEncoding:NSUTF8StringEncoding]); else lxwWs = workbook_add_worksheet(_lxwWorkbook, NULL);
    SIXLSXWorksheet* ws = [SIXLSXWorksheet _worksheetFromAdd:lxwWs];
    [_worksheets addObject:ws];
    return ws;
}

-(SIXLSXFormat*)addFormat {
    lxw_format* format = workbook_add_format(_lxwWorkbook);
    return [SIXLSXFormat _formatFromAdd:format];
}

-(SIXLSXChart*)addChartOfType:(SIXLSXChartType)chartType {
    lxw_chart* chart = workbook_add_chart(_lxwWorkbook, chartType);
    return [SIXLSXChart _chartFromAdd:chart];
}

-(BOOL)closeOrError:(NSError **)error {
    lxw_error lxwerr = workbook_close(_lxwWorkbook);
    if (lxwerr != 0) {
        *error = [NSError errorWithDomain:@"SIXLSXWriter" code:lxwerr userInfo:@{NSLocalizedDescriptionKey: [NSString stringWithUTF8String:lxw_strerror(lxwerr)]}];
        return NO;
    }
    return YES;
}

-(BOOL)setDocumentPropertiesFromDictionary:(NSDictionary *)propertiesDictionary {
    lxw_doc_properties props = {};
    for (NSString* key in propertiesDictionary.allKeys) {
        if ([[propertiesDictionary valueForKey:key] isKindOfClass:[NSString class]]) {
            NSString* tmpStr = [propertiesDictionary valueForKey:key];
            if ([key isEqualToString:@"title"]) props.title = (char*)[tmpStr cStringUsingEncoding:NSUTF8StringEncoding];
            else if ([key isEqualToString:@"subject"]) props.subject = (char*)[tmpStr cStringUsingEncoding:NSUTF8StringEncoding];
            else if ([key isEqualToString:@"author"]) props.author = (char*)[tmpStr cStringUsingEncoding:NSUTF8StringEncoding];
            else if ([key isEqualToString:@"manager"]) props.manager = (char*)[tmpStr cStringUsingEncoding:NSUTF8StringEncoding];
            else if ([key isEqualToString:@"company"]) props.company = (char*)[tmpStr cStringUsingEncoding:NSUTF8StringEncoding];
            else if ([key isEqualToString:@"category"]) props.category = (char*)[tmpStr cStringUsingEncoding:NSUTF8StringEncoding];
            else if ([key isEqualToString:@"keywords"]) props.keywords = (char*)[tmpStr cStringUsingEncoding:NSUTF8StringEncoding];
            else if ([key isEqualToString:@"comments"]) props.comments = (char*)[tmpStr cStringUsingEncoding:NSUTF8StringEncoding];
            else if ([key isEqualToString:@"hyperlink_base"]) props.hyperlink_base = (char*)[tmpStr cStringUsingEncoding:NSUTF8StringEncoding];
            else {
                if ([[propertiesDictionary valueForKey:key] isKindOfClass:[NSString class]]) {
                    lxw_error err = workbook_set_custom_property_string(_lxwWorkbook, [key cStringUsingEncoding:NSUTF8StringEncoding], (char*)[tmpStr cStringUsingEncoding:NSUTF8StringEncoding]);
                    if (err) return NO;
                }
            }
        } else {
            if ([[propertiesDictionary valueForKey:key] isKindOfClass:[NSNumber class]]) {
                lxw_error err = workbook_set_custom_property_number(_lxwWorkbook, [key cStringUsingEncoding:NSUTF8StringEncoding], [[propertiesDictionary valueForKey:key] doubleValue]);
                if (err) return NO;
            } else if ([[propertiesDictionary valueForKey:key] isKindOfClass:[NSDate class]]) {
                NSDate* date = [propertiesDictionary valueForKey:key];
                NSCalendar *calendar = [NSCalendar currentCalendar];
                NSDateComponents* datecomps = [calendar components:(NSCalendarUnitDay | NSCalendarUnitMonth | NSCalendarUnitYear | NSCalendarUnitHour | NSCalendarUnitMinute | NSCalendarUnitSecond) fromDate:date];
                lxw_datetime dt = {datecomps.year, datecomps.month, datecomps.day, datecomps.hour, datecomps.minute, (double)datecomps.second};
                lxw_error err = workbook_set_custom_property_datetime(_lxwWorkbook, [key cStringUsingEncoding:NSUTF8StringEncoding], &dt);
                if (err) return NO;
            }
        }
    }
    lxw_error err = workbook_set_properties(_lxwWorkbook, &props);
    if (err) return NO;
    return YES;
}

-(NSArray*)worksheets {
    return [NSArray arrayWithArray:_worksheets];
}

@end
