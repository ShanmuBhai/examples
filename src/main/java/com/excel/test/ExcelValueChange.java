package com.excel.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelValueChange {

	public static void main(String[] args) throws IOException {
		
		File file = new File("C:\\\\Users\\\\ELCOT\\\\eclipse-workspace\\\\FrameWork\\\\Excel\\\\Test demo.xlsx");
		
		FileInputStream stream = new FileInputStream(file);
		
		Workbook workbook = new XSSFWorkbook(stream);
		
		Sheet sheet = workbook.getSheet("Sheet1");
		
		Row row = sheet.getRow(1);
		
		Cell cell = row.getCell(8);
		
		String data = cell.getStringCellValue();
		
		if (data.equals("Android")) {
			
			//insert the value in file
			cell.setCellValue("Java");
		}
		FileOutputStream out = new FileOutputStream(file);
				
		workbook.write(out);
		
				
	}

}
