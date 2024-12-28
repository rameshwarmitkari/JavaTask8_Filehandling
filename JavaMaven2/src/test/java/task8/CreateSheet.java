package task8;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateSheet {

	public static void main(String[] args) throws IOException {
		XSSFWorkbook book = new XSSFWorkbook();
		XSSFSheet sheet = book.createSheet("Sheet1");

		String filepath = ".\\resource\\sheet.xlsx";

		Object data[][] = { { 1, "ram", "mitkari" }, 
				{ 23, "dec", "de" }, { 25, "qwec", "zxcde" } };


			FileOutputStream outputstream = new FileOutputStream(filepath);
			book.write(outputstream);
			outputstream.close();
			
		System.out.println("sheet1 is created");

	}

}
