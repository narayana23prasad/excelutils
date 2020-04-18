package utils;

import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excelutils {
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;
	
	//Constructor
	public excelutils(String excelpath,String sheetname) {
		try{
		workbook = new XSSFWorkbook(excelpath);
		sheet = workbook.getSheet(sheetname);
		}
		catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
	
	
	public static void main(String[] args) throws Exception {
		//getrowcount1();
		//getcelldata1();
	}
	
	public static void getcelldata1() throws IOException{
		String excelpath = "./data/testdata.xlsx";
		XSSFWorkbook workbook = new XSSFWorkbook(excelpath);
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		
		DataFormatter formatter = new DataFormatter();
		Object value = formatter.formatCellValue(sheet.getRow(1).getCell(2));
		
		String cellvalue1 = sheet.getRow(1).getCell(0).getStringCellValue();
		Double cellvalue2 = sheet.getRow(1).getCell(2).getNumericCellValue();
		System.out.println(cellvalue1);
		System.out.println(cellvalue2);
		System.out.println(value);
	}
	
	public static void getrowcount1(){
		try {
		String projdir = System.getProperty("user.dir");
		System.out.println(projdir);
		String excelpath = "./data/testdata.xlsx";
		XSSFWorkbook workbook = new XSSFWorkbook(excelpath);
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		int rowcount = sheet.getPhysicalNumberOfRows();
		System.out.println("No of rows: "+ rowcount);
		}
		
		catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	public static void getcelldata2(int getrownum,int getcolnum) throws IOException{
		
		DataFormatter formatter = new DataFormatter();
		Object value = formatter.formatCellValue(sheet.getRow(getrownum).getCell(getcolnum));
		System.out.println(value);
	}
	
	public static void getrowcount2(){
		try {

		int rowcount = sheet.getPhysicalNumberOfRows();
		System.out.println("No of rows: "+ rowcount);
		}
		
		catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	
	
}
