package com.ekam.utilities.excelwriter.impl;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import com.ekam.utilities.excelwriter.base.ExcelFileWriter;
import com.ekam.utilities.excelwriter.base.HeaderWriter;
import com.ekam.utilities.excelwriter.base.SheetDataWriter;
import com.ekam.utilities.excelwriter.base.WriteConfig;

public class GenericFileWriter implements ExcelFileWriter {
	private WriteConfig writeConfig;
	private HeaderWriter headerWriter;
	private SheetDataWriter sheetDataWriter;

	public GenericFileWriter(HeaderWriter headerWriter, SheetDataWriter sheetDataWriter) {
		this.headerWriter = headerWriter;
		this.sheetDataWriter = sheetDataWriter;
	}

	public void write(WriteConfig writeConfig) throws FileNotFoundException, IOException {
		checkPreconditions(writeConfig);
		this.writeConfig = writeConfig;
		initWriters(writeConfig);
		writeAllSheets(writeConfig);
		String outputLocation = writeConfig.getString(WriteConfig.CONFIG_OUTPUT_FILE_LOCATION);
		String fileName = writeConfig.getString(WriteConfig.CONFIG_OUTPUT_FILE_NAME);
		exportFileToLocation(outputLocation, fileName);
	}

	private void closeWriters(WriteConfig writeConfig) throws IOException {
		headerWriter.close();
		sheetDataWriter.close();
		writeConfig.get(WriteConfig.CONFIG_WORKBOOK, SXSSFWorkbook.class).close();
	}

	private void writeAllSheets(WriteConfig writeConfig) {
		int totalSheets = sheetDataWriter.totalSheetsToWrite();
		if (totalSheets > 0) {
			sheetDataWriter.beforeFileWritingStart();
			RowCellCounter rowCellCounter = new RowCellCounter();
			writeConfig.put(WriteConfig.CONFIG_CURRENT_ROW_CELL_COUNTER, rowCellCounter);
			for (int sheetCntr = 1; sheetCntr <= totalSheets; sheetCntr++) {
				Sheet sheet = writeConfig.get(WriteConfig.CONFIG_WORKBOOK, SXSSFWorkbook.class).createSheet();
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
		File outputLoc = new File(writeConfig.getString(WriteConfig.CONFIG_OUTPUT_FILE_LOCATION));
		outputLoc.mkdirs();
		File outputFile = new File(outputLoc, writeConfig.getString(WriteConfig.CONFIG_OUTPUT_FILE_NAME));
		outputFile.createNewFile();
		Workbook wb = new SXSSFWorkbook();
		writeConfig.put(WriteConfig.CONFIG_WORKBOOK, wb);
		headerWriter.init(writeConfig);
		sheetDataWriter.init(writeConfig);

	}

	protected void checkPreconditions(WriteConfig writeConfig) {
		checkNotNull(writeConfig, "Write Config parameter cannot be null");
		checkNotNull(writeConfig.getString(WriteConfig.CONFIG_OUTPUT_FILE_NAME),
				"Output File Name not provided in config");
		checkNotNull(writeConfig.getString(WriteConfig.CONFIG_OUTPUT_FILE_LOCATION),
				"Output File Path not provided in config");
	}

	private void checkNotNull(Object reference, String message) {
		if (reference == null) {
			throw new IllegalArgumentException(message);
		}
	}

	public final void close() throws IOException {
		closeWriters(writeConfig);
	}

	protected String exportFileToLocation(String exportLocation, String fileName) throws IOException {
		FileOutputStream out = null;
		File f = null;
		try {
			File directory = new File(exportLocation);
			if (!directory.exists()) {
				directory.mkdirs();
			}
			f = new File(exportLocation + File.separator + fileName);
			out = new FileOutputStream(f.getAbsolutePath());
			SXSSFWorkbook workbook = writeConfig.get(WriteConfig.CONFIG_WORKBOOK, SXSSFWorkbook.class);
			workbook.write(out);
			workbook.dispose();
			out.close();
		} finally {
			if (out != null) {
				out.close();
			}
		}
		return f != null ? f.getAbsolutePath() : "";
	}

}
