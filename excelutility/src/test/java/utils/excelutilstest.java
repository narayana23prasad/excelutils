package utils;

import java.io.IOException;

public class excelutilstest {
	public static void main(String[] args) throws IOException {
		String excelpath = "./data/testdata.xlsx";
		String sheetname = "Sheet1";
		excelutils excel = new excelutils(excelpath,sheetname);
		excel.getrowcount2();
		excel.getcelldata2(1,0);
		excel.getcelldata2(1,1);
		excel.getcelldata2(1,2);
	}
		

}
