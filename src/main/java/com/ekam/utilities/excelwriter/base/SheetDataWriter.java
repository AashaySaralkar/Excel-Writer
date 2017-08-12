package com.ekam.utilities.excelwriter.base;

public interface SheetDataWriter {
	
	void init(WriteConfig writerContext);
	int totalSheetsToWrite();
	void beforeFileWritingStart();
	void beforeSheetWritingStart();
	void writeSheetData(WriteConfig writeConfig);
	void afterSheetWritingEnd();
	void afterFileWritingEnd();
	void close();

	
}
