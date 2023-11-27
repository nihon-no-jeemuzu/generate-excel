/**
 * 
 */
package com.jemuzu.generate_excel.utility;

import java.io.BufferedOutputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Paths;
import java.util.concurrent.atomic.AtomicInteger;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.stereotype.Component;

import com.jemuzu.generate_excel.model.ExcelData;

import lombok.extern.slf4j.Slf4j;

@Slf4j
@Component
public class ExcelUtility {

	private ExcelData excelData;
	private String targetPath;

	private SXSSFWorkbook sworkbook;

	public ExcelUtility() {
	}

	public ExcelUtility(ExcelData excelData, String targetPath) {
		this.excelData = excelData;
		this.targetPath = Paths.get(targetPath, "generate_excel_output.xlsx").toString();
	}

	public void export() throws FileNotFoundException, IOException {
		try (BufferedOutputStream out = new BufferedOutputStream(new FileOutputStream(targetPath))) {
			log.debug("Creating file");
			writeData();
			sworkbook.write(out);
			out.close();
			log.debug("File Created");
		}
	}

	private void writeData() {
		AtomicInteger rowIndex = new AtomicInteger(1);
		AtomicInteger cellIndex = new AtomicInteger(0);
		try {
			sworkbook = new SXSSFWorkbook(); // add int to indicate number of rows to keep while writing data
			SXSSFSheet worksheet = sworkbook.createSheet("Data");

			log.debug("writing headers");
			Row headerRow = worksheet.createRow(0);
			cellIndex.set(0);
			excelData.getHeaders().forEach(header -> {
				Cell headerCell = headerRow.createCell(cellIndex.get());
				setCell(headerCell, header);
				worksheet.trackColumnForAutoSizing(cellIndex.getAndIncrement());
			});
			
			log.debug("writing data");
			excelData.getRows().forEach(line -> {
				Row dataRow = worksheet.createRow(rowIndex.getAndIncrement());
				cellIndex.set(0);
				line.forEach(data -> {
					Cell dataCell = dataRow.createCell(cellIndex.getAndIncrement());
					setCell(dataCell, data);
				});
				
			});
			
			log.debug("autosizing columns");
			cellIndex.set(0);
			excelData.getHeaders().forEach(header -> {
				worksheet.autoSizeColumn(cellIndex.getAndIncrement());
			});
			
		} catch (Exception e) {
			throw e;
		}
	}

	private void setCell(Cell cell, Object value) {
		if (value instanceof Integer) {
			cell.setCellValue((Integer) value);
		} else if (value instanceof Boolean) {
			cell.setCellValue((Boolean) value);
		} else {
			cell.setCellValue((String) value);
		}
	}

}