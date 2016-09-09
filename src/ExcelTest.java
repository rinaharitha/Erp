

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelTest {

	/**
	 * @param args
	 * @throws IOException 
	 */
	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		FileInputStream fi=new FileInputStream("d:\\demo.xlsx");
		String uid,pwd;
		
		XSSFWorkbook wb=new XSSFWorkbook(fi);
		XSSFSheet ws=wb.getSheet("LoginData");
		
		int rc=ws.getLastRowNum()+1;
		XSSFRow row;
		//uid=row.getCell(0).getStringCellValue();
		
		//pwd=row.getCell(1).getStringCellValue();
		/*for(int i=1; i<=rc; i++)
		{
			row=ws.getRow(i);
			uid=row.getCell(0).getStringCellValue();
			pwd=row.getCell(1).getStringCellValue();
			
		
		
	
		System.out.println(uid+" "+pwd);
		}*/
		
		row =ws.getRow(1);
		row.createCell(2).setCellValue("Pass");

		
		CellStyle passStyle=wb.createCellStyle();
		passStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
		passStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		row.getCell(2).setCellStyle(passStyle);
		FileOutputStream fo=new FileOutputStream("d:\\demo.xlsx");
		wb.write(fo);
	} 

}
