package com.ekam.utilities.excelwriter.impl;

public class RowCellCounter {
	
	private int rowNum;
	private int cellNum;
	
	public int nextRowNum(){
		cellNum = 0;
		return rowNum++;
	}
	
	public int nextCellNum(){
		return cellNum++;
	}
	
	public void reset(){
		rowNum = 0;
		cellNum = 0;
	}
	
	

}
