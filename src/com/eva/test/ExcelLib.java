package com.eva.test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;

public class ExcelLib {

	String filePath = "C:\\Users\\Araya.Kunal\\Desktop\\new.xls";

	/**
	 * 
	 * @param sheetNAme
	 *            , rowNum , colNum
	 * 
	 * @throws EncryptedDocumentException
	 * @throws InvalidFormatException
	 * @throws IOException
	 *             Read the data form ExcelSheet based user requirement return
	 *             always return string data
	 */
	public String getExcelData(String sheetNAme, int rowNum, int colNum)
			throws EncryptedDocumentException, InvalidFormatException,
			IOException {
		FileInputStream fis = new FileInputStream(filePath);
		Workbook wb = WorkbookFactory.create(fis);
		Sheet sh = wb.getSheet(sheetNAme);
		Row row = sh.getRow(rowNum);
		String data = row.getCell(colNum).getStringCellValue();
		return data;
	}

	

	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
		
		ExcelLib obj = new ExcelLib();
		String s = obj.getExcelData("kunal", 0, 0);
		String p = obj.getExcelData("kunal", 0, 1);
		String k = obj.getExcelData("kunal", 1, 0);
		String m = obj.getExcelData("kunal", 1, 1);
		
		
		System.out.println(s);
		System.out.println(p);
		System.out.println(k);
		System.out.println(m);
		
		WebDriver driver = new FirefoxDriver();
		driver.get("https://www.facebook.com");
		driver.findElement(By.id("email")).sendKeys(s);
		driver.findElement(By.id("pass")).sendKeys(p);
		
		
	}
	
}
