package com.ekam.utilities.excelwriter.impl;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import com.ekam.utilities.excelwriter.base.HeaderWriter;
import com.ekam.utilities.excelwriter.base.WriteConfig;

public abstract class ArrayBasedHeaderWriter implements HeaderWriter {

	public void write(WriteConfig writeConfig) {
		String[] headers = getHeaders();
		if (null != headers) {
			Sheet sheet = writeConfig.get(WriteConfig.CONFIG_CURRENT_SHEET,
					Sheet.class);
			RowCellCounter rowCellCounter = writeConfig.get(
					WriteConfig.CONFIG_CURRENT_ROW_CELL_COUNTER,
					RowCellCounter.class);
			Row headerRow = sheet.createRow(rowCellCounter.nextRowNum());
			for (String headerText : headers) {
				createHeaderCell(headerText, headerRow, rowCellCounter);
			}
		}

	}

	public abstract String[] getHeaders();

	public Cell createHeaderCell(String headerText, Row headerRow,
			RowCellCounter rowCellCounter) {
		Cell headerCell = headerRow.createCell(rowCellCounter.nextCellNum());
		headerCell.setCellValue(headerText);
		return headerCell;
	}

}
