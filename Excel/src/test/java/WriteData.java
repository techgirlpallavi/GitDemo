import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteData {

	public static void main(String[] args) throws IOException {
		//String path = "D:\\ExcelPractice\\WriteInto.xlsx";
		// FS has the file
		
		
		File file = new File("D:\\ExcelPractice\\WriteInto.xlsx");
		
		
		
		FileInputStream fis = new FileInputStream(file);
		XSSFWorkbook wb = new XSSFWorkbook(fis);

		XSSFSheet s1 = wb.getSheetAt(0);
		s1.getRow(0).getCell(1).setCellValue("Hello");
		int lastRow = s1.getLastRowNum();
		for (int i =1; i <= lastRow; i++) {
			Row row = s1.getRow(i);
			Cell cell = row.createCell(2);
			cell.setCellValue("hello hi");
		}
		
				
		FileOutputStream fos = new FileOutputStream(file);
		wb.write(fos);
		fos.close();
		System.out.println("Data is added to Excel File");
	}

}
