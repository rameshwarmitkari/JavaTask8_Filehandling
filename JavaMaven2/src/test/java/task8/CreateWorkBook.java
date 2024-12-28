package task8;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateWorkBook {

	public static void main(String[] args) throws Exception {

		XSSFWorkbook book = new XSSFWorkbook();
		XSSFSheet sheet = book.createSheet("java_Task");

		String filepath = ".\\resource\\task7.xlsx";

		Object data[][] = { { 1, "ram", "mitkari" }, 
				{ 23, "dec", "de" }, { 25, "qwec", "zxcde" } };

		int rows = data.length;
		int cols = data[0].length;

		for (int r = 0; r < rows; r++) {
			XSSFRow row = sheet.createRow(r);

			for (int c = 0; c < cols; c++) {

				XSSFCell cell = row.createCell(c);
				Object value = data[r][c];
				if (value instanceof String)
					cell.setCellValue((String) value);
				if (value instanceof Integer)
					cell.setCellValue((Integer) value);
				if (value instanceof Boolean)
					cell.setCellValue((Boolean) value);

			}
			FileOutputStream outputstream = new FileOutputStream(filepath);
			book.write(outputstream);
			outputstream.close();
		}
		System.out.println("task excel is created");

	}

}
