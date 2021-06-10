package com.Excel;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class CreateWorkbook {

	public static void main(String[] args) {

		// workbook
		HSSFWorkbook workbook = new HSSFWorkbook();
		// sheet
		HSSFSheet sheet = workbook.createSheet("Student list");

		// row head
		HSSFRow rowHead = sheet.createRow(0);
		rowHead.createCell(0).setCellValue("id");
		rowHead.createCell(1).setCellValue("StudentFName");
		rowHead.createCell(2).setCellValue("StudentLName");
		rowHead.createCell(3).setCellValue("USN");

		// first row
		HSSFRow row1 = sheet.createRow(1);

		row1.createCell(0).setCellValue(1);
		row1.createCell(1).setCellValue("Akshata");
		row1.createCell(2).setCellValue("Patil");
		row1.createCell(3).setCellValue("3GF16EC005");

		HSSFRow row2 = sheet.createRow(2);

		row2.createCell(0).setCellValue(2);
		row2.createCell(1).setCellValue("Sangeeta");
		row2.createCell(2).setCellValue("Hunje");
		row2.createCell(3).setCellValue("3GF16EC006");

		String path = "D:\\FileIO\\StudentDetails.xls";

		try {
			FileOutputStream fos = new FileOutputStream(path);
			workbook.write(fos);
			fos.close();
			workbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		System.out.println("changes made");

	}

}
