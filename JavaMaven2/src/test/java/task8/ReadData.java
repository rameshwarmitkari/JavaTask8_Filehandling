package task8;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadData {

	public static void main(String[] args) throws Exception {
		String filepath = ".\\resource\\data.xlsx";

		FileInputStream file1 = new FileInputStream(filepath);

		XSSFWorkbook book = new XSSFWorkbook(file1);

		XSSFSheet sheet = book.getSheetAt(0);

		int row = sheet.getLastRowNum();
		// System.out.println(sheet.getPhysicalNumberOfRows());
		System.out.println(row);
		int col = sheet.getRow(1).getLastCellNum();

		for (int i = 0; i < row; i++) {

			XSSFRow row1 = sheet.getRow(i);

			for (int j = 0; j < col; j++) {
				XSSFCell cell = row1.getCell(j);

				switch (cell.getCellType()) {
				case STRING:
					System.out.print(cell.getStringCellValue());
					break;
				case NUMERIC:
					System.out.print(cell.getNumericCellValue());
					break;
				case BOOLEAN:
					System.out.print(cell.getBooleanCellValue());
					break;
				}
				System.out.print(" | ");
			}
			System.out.println();

		}

	}
}
