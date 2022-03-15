package com.excel.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.commons.compress.archivers.dump.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

import io.github.bonigarcia.wdm.WebDriverManager;

import org.apache.poi.ss.usermodel.Row;

public class Student {

	public static void main(String[] args)
			throws IOException, org.apache.poi.openxml4j.exceptions.InvalidFormatException {

		WebDriverManager.chromedriver().setup();

		WebDriver driver = new ChromeDriver();
		driver.get("http://demo.automationtesting.in/Register.html");

		WebElement dDnSkill = driver.findElement(By.xpath("(//select[@type='text'])[1]"));

		Select select = new Select(dDnSkill);

		List<WebElement> options = select.getOptions();
		int len = options.size();
		System.out.println(len);

		for (int i = 0; i < options.size(); i++) {
			WebElement element = options.get(i);
			String text = element.getText();
			System.out.println(text);

			File file = new File("C:\\Users\\ELCOT\\eclipse-workspace\\FrameWork\\Excel\\Test demo.xlsx");

			FileInputStream stream = new FileInputStream(file);

			Workbook workbook = new XSSFWorkbook(stream);

			Sheet sheet = workbook.getSheet("Sheet2");

			Row row = sheet.createRow(i);

			Cell cell = row.createCell(0);

			cell.setCellValue(text);

			FileOutputStream out = new FileOutputStream(file);

			workbook.write(out);

		}

	}

}
