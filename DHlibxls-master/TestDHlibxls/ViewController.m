//
//  ViewController.m
//  TestDHlibxls
//
//  Created by David Hoerl on 1/10/12.
//  Copyright (c) 2012 __MyCompanyName__. All rights reserved.
//

#import "ViewController.h"

#import "DHxlsReader.h"

extern int xls_debug;

@implementation ViewController
{
	IBOutlet UITextView *textView;
}

- (void)didReceiveMemoryWarning
{
    [super didReceiveMemoryWarning];
    // Release any cached data, images, etc that aren't in use.
}

#pragma mark - View lifecycle

- (void)viewDidLoad
{
    [super viewDidLoad];

	NSString *path = [[[NSBundle mainBundle] resourcePath] stringByAppendingPathComponent:@"test1.xls"];
//	NSString *path = @"/tmp/test.xls";

	// xls_debug = 1; // good way to see everything in the Excel file
	
	DHxlsReader *reader = [DHxlsReader xlsReaderWithPath:path];
	assert(reader);

	NSString *text = @"";
	
	text = [text stringByAppendingFormat:@"AppName: %@\n", reader.appName];
	text = [text stringByAppendingFormat:@"Author: %@\n", reader.author];
	text = [text stringByAppendingFormat:@"Category: %@\n", reader.category];
	text = [text stringByAppendingFormat:@"Comment: %@\n", reader.comment];
	text = [text stringByAppendingFormat:@"Company: %@\n", reader.company];
	text = [text stringByAppendingFormat:@"Keywords: %@\n", reader.keywords];
	text = [text stringByAppendingFormat:@"LastAuthor: %@\n", reader.lastAuthor];
	text = [text stringByAppendingFormat:@"Manager: %@\n", reader.manager];
	text = [text stringByAppendingFormat:@"Subject: %@\n", reader.subject];
	text = [text stringByAppendingFormat:@"Title: %@\n", reader.title];
	
	text = [text stringByAppendingFormat:@"\n\nNumber of Sheets: %u\n", reader.numberOfSheets];

#if 0
	[reader startIterator:0];
	
	while(YES) {
		DHcell *cell = [reader nextCell];
		if(cell.type == cellBlank) break;
		
		text = [text stringByAppendingFormat:@"\n%@\n", [cell dump]];
	}
#else
    
    NSMutableDictionary *allDataResultDict = [[NSMutableDictionary alloc] initWithCapacity:100];
    
    NSMutableDictionary *cityResultDict = [[NSMutableDictionary alloc] initWithCapacity:100];
    
    //区别城市
    NSString *idStr = @"";
    
    //excel表中第几行开始
    int row = 2;
    
    while(YES) {
        
        //四列
        DHcell *cell1 = [reader cellInWorkSheetIndex:0 row:row col:1];
        DHcell *cell2 = [reader cellInWorkSheetIndex:0 row:row col:2];
        DHcell *cell3 = [reader cellInWorkSheetIndex:0 row:row col:3];
        DHcell *cell4 = [reader cellInWorkSheetIndex:0 row:row col:4];
        //如果某列为空，退出循环，否则崩溃
        if(cell1.type == cellBlank || cell2.type == cellBlank || cell3.type == cellBlank || cell4.type == cellBlank) break;
        
        NSMutableDictionary *provinceDict = nil;
        
        if (![idStr isEqualToString:cell1.dump]) {
            //如果是新城市，则进行存储(城市表)
            [cityResultDict setValue:cell2.dump forKey:cell1.dump];
            
            provinceDict = [[NSMutableDictionary alloc] initWithObjectsAndKeys:cell4.dump,cell3.dump, nil];
            
            idStr = cell1.dump;
            
        }else {
            
            provinceDict = [[NSMutableDictionary alloc] initWithDictionary:allDataResultDict[cell1.dump]];
            [provinceDict setValue:cell4.dump forKey:cell3.dump];
            
        }
        
        [allDataResultDict setValue:provinceDict forKey:cell1.dump];
        row++;
		
    }
    
    NSArray *paths = NSSearchPathForDirectoriesInDomains(NSDocumentDirectory, NSUserDomainMask, YES);
    NSString *plistPaht = [paths objectAtIndex:0];
    
    
    NSString *allDataResultDictFileName=[plistPaht stringByAppendingPathComponent:@"allDataResultDict.plist"];
    NSString *cityResultDictFileName=[plistPaht stringByAppendingPathComponent:@"cityResultDict.plist"];
    
    NSLog(@"fileName = %@",allDataResultDictFileName);

    //创建并写入文件
    [allDataResultDict writeToFile:allDataResultDictFileName atomically:YES];
    [cityResultDict writeToFile:cityResultDictFileName atomically:YES];
    
    //检查是否写入
    NSMutableDictionary *allDataResultWriteData=[[NSMutableDictionary alloc]initWithContentsOfFile:allDataResultDictFileName];
    NSMutableDictionary *cityResultDictData=[[NSMutableDictionary alloc]initWithContentsOfFile:cityResultDictFileName];
    NSLog(@"write data is :%@ %@",allDataResultWriteData,cityResultDictData);
    
#endif
	textView.text = text;
}

- (void)viewDidUnload
{
	textView = nil;
    [super viewDidUnload];
    // Release any retained subviews of the main view.
    // e.g. self.myOutlet = nil;
}

- (void)viewWillAppear:(BOOL)animated
{
    [super viewWillAppear:animated];
}

- (void)viewDidAppear:(BOOL)animated
{
    [super viewDidAppear:animated];
}

- (void)viewWillDisappear:(BOOL)animated
{
	[super viewWillDisappear:animated];
}

- (void)viewDidDisappear:(BOOL)animated
{
	[super viewDidDisappear:animated];
}

- (BOOL)shouldAutorotateToInterfaceOrientation:(UIInterfaceOrientation)interfaceOrientation
{
    // Return YES for supported orientations
	return (interfaceOrientation != UIInterfaceOrientationPortraitUpsideDown);
}

@end
