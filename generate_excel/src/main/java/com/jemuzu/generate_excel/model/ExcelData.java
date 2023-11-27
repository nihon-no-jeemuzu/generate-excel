package com.jemuzu.generate_excel.model;

import java.util.List;

import lombok.Data;

@Data
public class ExcelData {
	
	private List<String> headers;
	private List<List<String>> rows;

}
