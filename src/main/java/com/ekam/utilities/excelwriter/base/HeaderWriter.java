package com.ekam.utilities.excelwriter.base;

public interface HeaderWriter {
	
	void init(WriteConfig writeConfig);
	void write(WriteConfig writeConfig);
	void close(); 

}
