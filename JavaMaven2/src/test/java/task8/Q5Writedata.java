package task8;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Q5Writedata {

	public static void main(String[] args) throws Exception {
		XSSFWorkbook book = new XSSFWorkbook();
		XSSFSheet sheet = book.createSheet("Q5");

		String filepath = ".\\resource\\Task8_Q5.xlsx";

		Object data[][] = { { "sophie", 30, "john@test.com" }, 
				{ "martin", 28, "john@test.com" }, 
					{ "steve", 35, "jacky@example.com" },
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
		System.out.println("Data has been written in sheet Q5");



	}

}
