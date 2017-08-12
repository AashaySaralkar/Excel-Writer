package com.ekam.utilities.excelwriter.base;

import java.util.HashMap;
import java.util.Map;

public final class WriteConfig {
	
	public static final String CONFIG_OUTPUT_FILE_LOCATION = "CONFIG_FILE_LOCATION";
	public static final String CONFIG_OUTPUT_FILE_NAME = "CONFIG_OUTPUT_FILE_NAME";
	public static final String CONFIG_CURRENT_SHEET ="CONFIG_CURRENT_SHEET";
	public static final String CONFIG_WORKBOOK = "CONFIG_WORKBOOK";
	public static final String CONFIG_CURRENT_ROW_CELL_COUNTER ="CONFIG_CURRENT_ROW_CELL_COUNTER";
	
	private Map<String, Object> data;

	public WriteConfig() {
		data = new HashMap<String, Object>();
	}

	public WriteConfig put(String key, Object value) {
		data.put(key, value);
		return this;
	}
	
	public String getString(String key){
		return (String)data.get(key);
	}
	
	public <T> T get(String key, Class<T> type){
		return type.cast(data.get(key));
	}

	public Object remove(String key) {
		return data.remove(key);
	}
	
	public boolean containsKey(String key){
		return data.containsKey(key);
	}

}
