package com.excel.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Employee {
	@SuppressWarnings("deprecation")
	public static void main(String[] args) throws IOException {

		// Mention the path of the excel
		File file = new File("C:\\Users\\ELCOT\\eclipse-workspace\\FrameWork\\Excel\\Test demo.xlsx");

		// Get the objects or bytes from file

		FileInputStream Stream = new FileInputStream(file);

		// create the workbook

		Workbook workbook = new XSSFWorkbook(Stream);

		// Get the Sheetbname
		Sheet sheet = workbook.getSheet("Sheet1");

		// get the all row & cell

		for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
			Row row = sheet.getRow(i);

			for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
				Cell cell = row.getCell(j);
				//System.out.println(cell);
				// cell type string -->text numeric-->Number
				CellType cellType = cell.getCellType();
				// condition for checking string and number
				switch (cellType) {
				case STRING:
					String data = cell.getStringCellValue();
					System.out.println(data);
					break;
				case NUMERIC:
					if (DateUtil.isCellDateFormatted(cell)) {
						String dataInfo = new SimpleDateFormat("dd-MMM-YY").format(cell.getDateCellValue());
						System.out.println(dataInfo);

					} else {
						String name = BigDecimal.valueOf(cell.getNumericCellValue()).toString();
						System.out.println(name);

					}
					break;
				default:
					break;
				}

			}

		}

	}
}
