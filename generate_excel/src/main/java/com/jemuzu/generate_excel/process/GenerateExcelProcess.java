package com.jemuzu.generate_excel.process;

import java.nio.file.Files;
import java.nio.file.LinkOption;
import java.nio.file.Paths;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import com.jemuzu.generate_excel.model.ExcelData;
import com.jemuzu.generate_excel.utility.ExcelUtility;
import com.jemuzu.generate_excel.utility.ParseData;

import lombok.extern.slf4j.Slf4j;

@Slf4j
@Component
public class GenerateExcelProcess {

	@Autowired
	ParseData parseData;

	public void executeProcess(String sourcePath, String targetPath) throws Exception {

		if (Files.isDirectory(Paths.get(targetPath), LinkOption.NOFOLLOW_LINKS)) {
			ExcelData excelData = parseData.parseSourceData(sourcePath);
			ExcelUtility excelUtility = new ExcelUtility(excelData, targetPath);
			excelUtility.export();
		} else {
			log.error("ERROR: targetPath does not exists.");
			System.exit(1);
		}
	}

}
