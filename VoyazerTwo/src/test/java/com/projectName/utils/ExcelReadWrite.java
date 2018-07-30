package test.java.com.projectName.utils;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReadWrite {
	public static String filename = System.getProperty("user.dir") + "\\src\\config\\testcases\\TestData.xlsx";
	public String path;
	public FileInputStream fis = null;
	public FileOutputStream fileOut = null;
	private XSSFWorkbook workbook = null;
	private XSSFSheet sheet = null;
	private XSSFRow row = null;
	private XSSFCell cell = null;

	public void ReadExcel(String fpath) {

		this.path = path;
		try {
			fis = new FileInputStream(fpath);
			workbook = new XSSFWorkbook();
			sheet = workbook.getSheetAt(0);
			fis.close();
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	public boolean CreateSheet(String fname) {

		this.fileOut = fileOut;

		try {

			workbook = new XSSFWorkbook();
			workbook.createSheet();
			fileOut.close();
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}

	/*public string getcelldata(String fname,String colName, int rownum) {
		this.fis = fis;
		workbook = new XSSFWorkbook();	
		
		if(colName.isEmpty() || fname.isEmpty() || rowNum<=0)
			return "";			
		try {
			int index = workbook.getSheetIndex(fname);
			if (index<=0) return "";
			sheet = workbook.getSheetAt(index);
			row = sheet.getRow(rownum-1);
			if(row==null)
				return "";
			cell = row.getCell();
						
			
		}catch{
			
		}
		
		return rowNum;
		
		
	}
}*/
