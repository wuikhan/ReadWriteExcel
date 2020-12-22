package readWriteExcelPackage;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadWriteExcel {
	public static FileOutputStream fos;
	public static FileInputStream fis;
	public static XSSFWorkbook wb;
	public static XSSFSheet sh;
	public static int rowcount;
	public static int cellcount;
	public static String datafile;
	public static String cellValue;
	public static XSSFCell cell;

	public static void main(String[] args) throws IOException {
		datafile = "C:\\Users\\waqas.khan\\eclipse-workspace\\ReadWriteExcel\\file.xlsx";
		fis = new FileInputStream(datafile);
		wb = new XSSFWorkbook(fis);
		sh = wb.getSheetAt(0);

		rowcount = sh.getLastRowNum();
		for (int i = 1; i <= rowcount; i++) {
			cellcount = sh.getRow(i).getLastCellNum();
			for (int j = 0; j < cellcount; j++) {
				cellValue = sh.getRow(i).getCell(j).getStringCellValue();
				fos = new FileOutputStream(datafile);
				cell = wb.getSheetAt(0).getRow(i).createCell(2);
				if (cellValue.contains("pa55word")) {
					cell.setCellValue("Pass");
				} else {
					cell.setCellValue("Fail");
				}
			}
		}
		wb.write(fos);
		wb.close();
		System.out.println("reading & writing complete");
	}
}