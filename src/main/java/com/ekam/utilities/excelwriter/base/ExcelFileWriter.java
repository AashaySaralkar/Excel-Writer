package com.ekam.utilities.excelwriter.base;

import java.io.FileNotFoundException;
import java.io.IOException;

public interface ExcelFileWriter {

	public void write(WriteConfig writeConfig) throws FileNotFoundException, IOException;

	public void close();

}
