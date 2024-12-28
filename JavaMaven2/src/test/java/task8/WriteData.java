package task8;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteData {

	public static void main(String[] args) throws Exception  {


		XSSFWorkbook book = new XSSFWorkbook();
		XSSFSheet sheet = book.createSheet("Q3");

		String filepath = ".\\resource\\Task8_Q3.xlsx";

		Object data[][] = { { "John Doe", 30, "john@test.com" }, 
				{ "John Doe", 28, "john@test.com" }, 
					{ "Bob Smith", 35, "jacky@example.com" },
					{ "swapnil", 37, "swapnil@example.com" } };

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
		System.out.println("Data has been written in sheet");


	}

}
