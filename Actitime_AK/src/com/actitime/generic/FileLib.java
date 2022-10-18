package com.actitime.generic;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Properties;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * this generic class is used to read the data from files such as property files, Excel files.
 * @author Akash RJ
 *
 */

public class FileLib {
	
	/**
	 * this method is used to get data from the property file
	 * @param key
	 * @return
	 * @throws IOException
	 */
	
	public String getPropertyData(String key) throws IOException
	{
		FileInputStream fis = new FileInputStream("./data/commondata.property");
		Properties p = new Properties();
		p.load(fis);
		String data = p.getProperty(key);
		return data;
	}
	
	/**
	 * this method is used to get data from excel sheet
	 * @param sheetname
	 * @param row
	 * @param cell
	 * @return
	 * @throws EncryptedDocumentException
	 * @throws IOException
	 */
	public String getExcelData(String sheetname, int row, int cell) throws EncryptedDocumentException, IOException
	{
		FileInputStream fis = new FileInputStream("./data/Book1.xlsx");
		Workbook wb = WorkbookFactory.create(fis);
		String data = wb.getSheet(sheetname).getRow(row).getCell(cell).getStringCellValue();
		return data;
		
		
	}
	
	/**
	 * this method is used to update the data in excel
	 * @param sheetname
	 * @param row
	 * @param cell
	 * @param setvalue
	 * @throws EncryptedDocumentException
	 * @throws IOException
	 */
	
	public void setExcelData(String sheetname, int row, int cell,String setvalue) throws EncryptedDocumentException, IOException
	{
		FileInputStream fis = new FileInputStream("./data/Book1.xlsx");
		Workbook wb = WorkbookFactory.create(fis);
		wb.getSheet(sheetname).getRow(row).getCell(cell).setCellValue(setvalue);
		
		FileOutputStream fos = new FileOutputStream("./data/Book1.xlsx");
		wb.write(fos);
		wb.close();
		
		
		
	}

}
