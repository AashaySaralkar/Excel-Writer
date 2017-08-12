package com.ekam.utilities.excelwriter.impl;

import org.junit.Test;

import com.ekam.utilities.excelwriter.base.HeaderWriter;
import com.ekam.utilities.excelwriter.base.SheetDataWriter;
import com.ekam.utilities.excelwriter.base.WriteConfig;

public class GenericWriterTests {
	
	@Test(expected=IllegalArgumentException.class)
	public void throwsExceptionIfWriteConfigIsEmpty() throws Exception{
		
		TestHeaderWriter headerWriter = new TestHeaderWriter();
		TestDataWriter dataWriter = new TestDataWriter();
		GenericFileWriter genericFileWriter = new GenericFileWriter(headerWriter,dataWriter);
		WriteConfig config = new WriteConfig();
		genericFileWriter.write(config);
		genericFileWriter.close();
	}
	
	@Test(expected=IllegalArgumentException.class)
	public void throwsExceptionIfFileLocationIsEmpty() throws Exception{
		
		TestHeaderWriter headerWriter = new TestHeaderWriter();
		TestDataWriter dataWriter = new TestDataWriter();
		GenericFileWriter genericFileWriter = new GenericFileWriter(headerWriter,dataWriter);
		WriteConfig config = new WriteConfig();
		genericFileWriter.write(config);
		genericFileWriter.close();
	}
	
	
	
	
	
	class TestHeaderWriter implements  HeaderWriter{

		public void init(WriteConfig writeConfig) {
			
		}

		public void write(WriteConfig writeConfig) {
			
		}

		public void close() {
			
		}
		
	}
	
	class TestDataWriter implements SheetDataWriter{

		public void init(WriteConfig writerContext) {
			// TODO Auto-generated method stub
			
		}

		public int totalSheetsToWrite() {
			// TODO Auto-generated method stub
			return 0;
		}

		public void beforeFileWritingStart() {
			// TODO Auto-generated method stub
			
		}

		public void beforeSheetWritingStart() {
			// TODO Auto-generated method stub
			
		}

		public void writeSheetData(WriteConfig writeConfig) {
			// TODO Auto-generated method stub
			
		}

		public void afterSheetWritingEnd() {
			// TODO Auto-generated method stub
			
		}

		public void afterFileWritingEnd() {
			// TODO Auto-generated method stub
			
		}

		public void close() {
			// TODO Auto-generated method stub
			
		}
		
	}

}
