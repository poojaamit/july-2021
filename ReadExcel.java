package demo;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class ReadExcel {
@Test
public void readdatafromexcekl() throws IOException {
	FileInputStream file = new FileInputStream("C:\\Users\\AMIT\\Downloads\\user_detail_2020810.xlsx");
	XSSFWorkbook workbook = new XSSFWorkbook(file);
	XSSFSheet sheet = workbook.getSheetAt(0);
	
	System.out.println(sheet.getRow(0).getCell(0).getStringCellValue());
	System.out.println(sheet.getRow(1).getCell(0).getNumericCellValue());
	
	Row row = sheet.createRow(2);
	Cell cell = row.createCell(5);
	cell.setCellValue("ops pvt lmt");
	FileOutputStream fil = new FileOutputStream("E:\\Pooja\\Book1.xlsx");
	workbook.write(fil);;
	file.close();
	System.out.println("end of writting data in excel");
	
}
}
