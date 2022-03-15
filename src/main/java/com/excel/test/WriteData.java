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

public class WriteData {

	private void writeData(String sheetName, int rowNum, int cellNum, String data) throws IOException {

		File file = new File("C:\\Users\\ELCOT\\eclipse-workspace\\FrameWork\\Excel\\Test demo.xlsx");

		FileInputStream stream = new FileInputStream(file);

		Workbook workbook = new XSSFWorkbook(stream);

		Sheet sheet = workbook.getSheet(sheetName);

		Row row = sheet.getRow(rowNum);

		Cell cell = row.createCell(cellNum);

		cell.setCellValue(data);

		FileOutputStream outputStream = new FileOutputStream(file);

		workbook.write(outputStream);

	}

}
