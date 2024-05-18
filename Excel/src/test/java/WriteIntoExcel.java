import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteIntoExcel {

	public static void main(String[] args) {
		// TODO Auto-generated method stub

	File file = new File("D:\\\\ExcelPractice\\WriteInto.xlsx");
	XSSFWorkbook wb = new XSSFWorkbook();
	XSSFSheet sh = wb.createSheet();
	sh.createRow(0).createCell(0).setCellValue("Age");
	try {
		
		FileOutputStream fos = new FileOutputStream(file);
		wb.write(fos);
	}
	catch(Exception e)
	{
		e.printStackTrace();
	}
	
	
	}

}
