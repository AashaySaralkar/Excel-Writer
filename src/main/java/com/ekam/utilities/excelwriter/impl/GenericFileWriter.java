package com.ekam.utilities.excelwriter.impl;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.ekam.utilities.excelwriter.base.ExcelFileWriter;
import com.ekam.utilities.excelwriter.base.HeaderWriter;
import com.ekam.utilities.excelwriter.base.SheetDataWriter;
import com.ekam.utilities.excelwriter.base.WriteConfig;

public class GenericFileWriter implements ExcelFileWriter {
	private HeaderWriter headerWriter;
	private SheetDataWriter sheetDataWriter;

	public GenericFileWriter(HeaderWriter headerWriter,
			SheetDataWriter sheetDataWriter) {
		this.headerWriter = headerWriter;
		this.sheetDataWriter = sheetDataWriter;
	}

	public void write(WriteConfig writeConfig) throws FileNotFoundException, IOException {
		checkPreconditions(writeConfig);
		initWriters(writeConfig);
		writeAllSheets(writeConfig);
		closeWriters(writeConfig);
	}

	private void closeWriters(WriteConfig writeConfig) throws IOException {
		headerWriter.close();
		sheetDataWriter.close();
		writeConfig.get(WriteConfig.CONFIG_WORKBOOK, Workbook.class).close();
	}

	private void writeAllSheets(WriteConfig writeConfig) {
		int totalSheets = sheetDataWriter.totalSheetsToWrite();
		if(totalSheets > 0){
			sheetDataWriter.beforeFileWritingStart();
			RowCellCounter rowCellCounter =  new RowCellCounter();
			writeConfig.put(WriteConfig.CONFIG_CURRENT_ROW_CELL_COUNTER, rowCellCounter);
			for (int sheetCntr = 1; sheetCntr <= totalSheets; sheetCntr++){
				Sheet sheet = writeConfig.get(WriteConfig.CONFIG_WORKBOOK, Workbook.class).createSheet();
				writeConfig.put(WriteConfig.CONFIG_CURRENT_SHEET, sheet);
				sheetDataWriter.beforeSheetWritingStart();
				headerWriter.write(writeConfig);
				sheetDataWriter.writeSheetData(writeConfig);
				sheetDataWriter.afterSheetWritingEnd();
				writeConfig.get(WriteConfig.CONFIG_CURRENT_ROW_CELL_COUNTER, RowCellCounter.class).reset();
			}
			sheetDataWriter.afterFileWritingEnd();
		}
	}

	private void initWriters(WriteConfig writeConfig) throws FileNotFoundException, IOException {
		headerWriter.init(writeConfig);
		sheetDataWriter.init(writeConfig);
		File outputLoc = new File(writeConfig.getString(WriteConfig.CONFIG_OUTPUT_FILE_LOCATION));
		outputLoc.mkdirs();
		File outputFile = new File(outputLoc,writeConfig.getString(WriteConfig.CONFIG_OUTPUT_FILE_NAME));
		Workbook wb = new HSSFWorkbook(new FileInputStream(outputFile));
		writeConfig.put(WriteConfig.CONFIG_WORKBOOK, wb);
		
	}

	protected void checkPreconditions(WriteConfig writeConfig) {
		checkNotNull(writeConfig, "Write Config parameter cannot be null");
		checkNotNull(writeConfig.getString(WriteConfig.CONFIG_OUTPUT_FILE_NAME),"Output File Name not provided in config");
		checkNotNull(writeConfig.getString(WriteConfig.CONFIG_OUTPUT_FILE_LOCATION),"Output File Path not provided in config");
	}
	
	private void checkNotNull(Object reference,String message){
		if(reference == null){
			throw new IllegalArgumentException(message);
		}
	}

	public final void close() {
		close();
	}
}
