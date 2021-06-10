package com.Excel;

import java.io.FileInputStream;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class ReadData {

	public static void main(String[] args) {
		String path = "D:\\FileIO\\StudentDetails.xls";

		try {
			FileInputStream fis = new FileInputStream(path);

			// create workbook for storing
			HSSFWorkbook workBook = new HSSFWorkbook(fis);

			HSSFSheet sheet = workBook.getSheetAt(0);

		
			System.out.println(sheet.getRow(1).getCell(2));
			System.out.println("================================");
			for (Row row : sheet) {
				for (Cell cell : row) {
					switch (cell.getCellType()) {
					case NUMERIC:
						System.out.print((int) cell.getNumericCellValue() + " ");
						break;
					case STRING:
						System.out.print(cell.getStringCellValue() + " ");
						break;
					default:
						break;
					}

				}
				System.out.println();
			}
			fis.close();
			workBook.close();
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}
}
