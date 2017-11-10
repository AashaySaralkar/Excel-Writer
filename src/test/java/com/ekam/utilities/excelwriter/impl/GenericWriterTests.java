package com.ekam.utilities.excelwriter.impl;

import java.io.InputStreamReader;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.junit.Assert;
import org.junit.Test;

import com.ekam.utilities.excelwriter.base.SheetDataWriter;
import com.ekam.utilities.excelwriter.base.WriteConfig;
import com.google.gson.Gson;
import com.google.gson.stream.JsonReader;

public class GenericWriterTests {

	private static final String[] HEADERS = "Name,Type, Continent, Size,Status".split(",");
	@Test
	public void test() {
		try {
			TestHeaderWriter headerWriter = new TestHeaderWriter();
			
			TestDataWriter dataWriter = new TestDataWriter();
			GenericFileWriter genericFileWriter = new GenericFileWriter(headerWriter,dataWriter);
			WriteConfig config = new WriteConfig();
			config.put(WriteConfig.CONFIG_OUTPUT_FILE_NAME, "airports.xlsx");
			config.put(WriteConfig.CONFIG_OUTPUT_FILE_LOCATION, "D:\\ExcelWriter");
			genericFileWriter.write(config);
			genericFileWriter.close();
		} catch (Exception e) {
			Assert.fail("Excel generation failed due to "+e);
		}
	}
	
	
	
	class TestHeaderWriter extends ArrayBasedHeaderWriter{

		public void init(WriteConfig writeConfig) {
			// Nothing as of now
			
		}

		public void close() {
			//Nothing as of now
			
		}

		@Override
		public String[] getHeaders() {
			return HEADERS;
		}

		
		
	}
	
	class TestDataWriter implements SheetDataWriter{
		
		JsonReader reader;
		Gson gson;
		CellStyle airportStyle;
		CellStyle helicoptorStyle;

		public void init(WriteConfig writerContext){
			reader = new JsonReader(new InputStreamReader(getClass().getClassLoader().getResourceAsStream("airports.json")));
			reader.setLenient(Boolean.TRUE);
			gson = new Gson();
			Workbook workbook = writerContext.get(WriteConfig.CONFIG_WORKBOOK, Workbook.class);
	        Font headerFont = workbook.createFont();
	        headerFont.setBold(true);
	        airportStyle = createBorderedStyle(workbook);
	        airportStyle.setAlignment(HorizontalAlignment.CENTER);
	        airportStyle.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
	        airportStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	        airportStyle.setFont(headerFont);

	        helicoptorStyle = createBorderedStyle(workbook);
	        helicoptorStyle.setAlignment(HorizontalAlignment.CENTER);
	        helicoptorStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
	        helicoptorStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	        helicoptorStyle.setFont(headerFont);
			
		}

		public int totalSheetsToWrite() {
			return 1;
		}

		public void beforeFileWritingStart() {
			
			
		}

		public void beforeSheetWritingStart() {
			// Nothing as of now
			
		}

		public void writeSheetData(WriteConfig writeConfig) {
			try {
				reader.beginArray();
				SXSSFSheet sheet = writeConfig.get(WriteConfig.CONFIG_CURRENT_SHEET, SXSSFSheet.class);
				RowCellCounter rowCellCounter = writeConfig.get(WriteConfig.CONFIG_CURRENT_ROW_CELL_COUNTER, RowCellCounter.class);
				sheet.trackColumnForAutoSizing(0);
				sheet.trackColumnForAutoSizing(1);
				while (reader.hasNext()) {
					Airport airport = gson.fromJson(reader, Airport.class);
					if(StringUtils.isNoneEmpty(airport.getName())){
						Row row = sheet.createRow(rowCellCounter.nextRowNum());
						
						Cell cell = row.createCell(rowCellCounter.nextCellNum());
						cell.setCellValue(airport.getName());
						cell.setCellStyle("airport".equals(airport.getType()) ? airportStyle : helicoptorStyle);
						
						cell = row.createCell(rowCellCounter.nextCellNum());
						cell.setCellValue(airport.getType());			
						
						cell = row.createCell(rowCellCounter.nextCellNum());
						cell.setCellValue(airport.getContinent());	
						
						cell = row.createCell(rowCellCounter.nextCellNum());
						cell.setCellValue(StringUtils.isEmpty(airport.getSize()) ? "NA":airport.getSize());					
						
						cell = row.createCell(rowCellCounter.nextCellNum());
						cell.setCellValue(StringUtils.equals(airport.getStatus(),"1") ? "Active" : "Inactive");
						
					}
				}
				sheet.autoSizeColumn(0);
				sheet.autoSizeColumn(1);
				reader.endArray();
				reader.close();
			} catch (Exception e) {
				e.printStackTrace();
			}
			
		}

		public void afterSheetWritingEnd() {
			// Nothing as of now
			
		}

		public void afterFileWritingEnd() {
			// Nothing as of now
			
		}

		public void close() {
			// Nothing as of now
		}
		
		private CellStyle createBorderedStyle(Workbook wb){
	        BorderStyle thin = BorderStyle.THIN;
	        short black = IndexedColors.BLACK.getIndex();
	        
	        CellStyle style = wb.createCellStyle();
	        style.setBorderRight(thin);
	        style.setRightBorderColor(black);
	        style.setBorderBottom(thin);
	        style.setBottomBorderColor(black);
	        style.setBorderLeft(thin);
	        style.setLeftBorderColor(black);
	        style.setBorderTop(thin);
	        style.setTopBorderColor(black);
	        return style;
	    }
		
	}
	
	@SuppressWarnings("unused")
	private class Airport{
		private String name;
		private String type;
		private String continent;
		private String size;
		private String status;
		/**
		 * @return the name
		 */
		public final String getName() {
			return name;
		}
		/**
		 * @param name the name to set
		 */
		public final void setName(String name) {
			this.name = name;
		}
		/**
		 * @return the type
		 */
		public final String getType() {
			return type;
		}
		/**
		 * @param type the type to set
		 */
		public final void setType(String type) {
			this.type = type;
		}
		/**
		 * @return the continent
		 */
		public final String getContinent() {
			return continent;
		}
		/**
		 * @param continent the continent to set
		 */
		public final void setContinent(String continent) {
			this.continent = continent;
		}
		/**
		 * @return the size
		 */
		public final String getSize() {
			return size;
		}
		/**
		 * @param size the size to set
		 */
		public final void setSize(String size) {
			this.size = size;
		}
		/**
		 * @return the status
		 */
		public final String getStatus() {
			return status;
		}
		/**
		 * @param status the status to set
		 */
		public final void setStatus(String status) {
			this.status = status;
		}
		
	}

}


