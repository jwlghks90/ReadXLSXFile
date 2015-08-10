package ReadXLSXFile;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadXLSXFile {
	public ReadXLSXFile(){}
	
	public void readXLSX(){
		try {
			InputStream excelFileToRead = new FileInputStream("C:/test.xlsx"); // xlsx file path 
			XSSFWorkbook workbook = new XSSFWorkbook(excelFileToRead);
			
			XSSFSheet sheet = workbook.getSheetAt(0);
			XSSFRow row; 
			XSSFCell cell;

			Iterator rows = sheet.rowIterator();
			String stringData;
			int integerData;
			
			while (rows.hasNext()){
				
				row=(XSSFRow) rows.next();
				Iterator cells = row.cellIterator();
				
				cell = (XSSFCell) cells.next();
				
				if (cell.getCellType() == XSSFCell.CELL_TYPE_STRING){ // String
					stringData = cell.getStringCellValue();
					System.out.println("data : " + stringData);
				}
				else if(cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC){ // int
					integerData = (int)cell.getNumericCellValue();
					System.out.println("data : " + integerData);
				}
				
			}
			
		}catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
