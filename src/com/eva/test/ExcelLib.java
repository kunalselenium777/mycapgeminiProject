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
	
	public int getRowCount(String sheetNAme) throws EncryptedDocumentException, 
	InvalidFormatException, IOException{
			  FileInputStream fis = new FileInputStream(filePath);
			    Workbook wb = WorkbookFactory.create(fis);
			    Sheet sh = wb.getSheet(sheetNAme);
			    int rowNum = sh.getLastRowNum();
			return rowNum;		
		}
	public void setExcelData(String sheetNAme, int rowNum , 
			int colNum, String data) throws EncryptedDocumentException, 
			InvalidFormatException, IOException{
					  FileInputStream fis = new FileInputStream(filePath);
					    Workbook wb = WorkbookFactory.create(fis);
					    Sheet sh = wb.getSheet(sheetNAme);
					    Row row = sh.getRow(rowNum);
					    FileOutputStream fos = new FileOutputStream(filePath);
					    Cell cel = row.createCell(colNum);
					    cel.setCellValue(data);
					    wb.write(fos);
					    wb.close();
					
				}



	public static void main(String[] args) throws EncryptedDocumentException,
			InvalidFormatException, IOException {

		
		ExcelLib obj = new ExcelLib();
		String userName = obj.getExcelData("kunal", 0, 0);
		String passWord = obj.getExcelData("kunal", 1, 0);
		String Address = obj.getExcelData("kunal", 2, 0);
		String phoneNumber = obj.getExcelData("kunal", 0, 1);
		String rollNumber = obj.getExcelData("kunal",1,1);
		String salary = obj.getExcelData("kunal", 2, 1);
		
		
		System.out.println(userName);
		System.out.println(passWord);
		System.out.println(Address);
		System.out.println(phoneNumber);
		System.out.println(rollNumber);
		System.out.println(salary);
		

//		WebDriver driver = new FirefoxDriver();
//		driver.get("https://www.facebook.com");
//		driver.findElement(By.id("email")).sendKeys(userName);
//		driver.findElement(By.id("pass")).sendKeys(passWord);

	}

}
