package com.excel.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class UpdateData {

	public void Getdata(String sheetName, int rowNum, int cellNum, String olddata, String newdata) throws IOException {

		File file = new File("C:\\Users\\ELCOT\\eclipse-workspace\\FrameWork\\Excel\\Test demo.xlsx");

		FileInputStream stream = new FileInputStream(file);

		Workbook workbook = new XSSFWorkbook(stream);

		Sheet sheet = workbook.getSheet(sheetName);

		Row row = sheet.getRow(rowNum);

		Cell cell = row.getCell(cellNum);

		String value = cell.getStringCellValue();

		if (value.equals(olddata)) {
			cell.setCellValue(newdata);

		}

		FileOutputStream outputStream = new FileOutputStream(file);

		workbook.write(outputStream);

	}
}
