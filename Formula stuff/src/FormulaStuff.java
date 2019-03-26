import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;

import javax.swing.JButton;
import javax.swing.JFrame;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.IndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellFill;

public class FormulaStuff {
	public static void main(String[] args) {
		//creating some variables
		String binaryString;
		File f;
		FileInputStream fis;
		Workbook wb;
		String b = "1";
		try{
			//Opening workbook
			f = new File("C:\\Users\\samuel\\Documents\\Tupper's Formula\\Editable Formula2.xlsx");
			fis = new FileInputStream(f);
			wb = new XSSFWorkbook(fis);

			//getting the binary value
			XSSFSheet sheet1 = (XSSFSheet) wb.getSheetAt(1);
			XSSFSheet sheet2 = (XSSFSheet) wb.getSheetAt(0);
			b = sheet1.getRow(5).getCell(2).getStringCellValue();
			
			//Converting binary string to a very large num
			binaryString = b.replace("	", "");
			BigInteger bs = new BigInteger(binaryString,2);
			bs = bs.multiply(new BigInteger("17"));
			
			//Format the cell to fit want format
			XSSFCell cell = sheet2.createRow(25).createCell(2);
			XSSFCellStyle cs = (XSSFCellStyle) wb.getCellStyleAt(1);
			cs.setWrapText(true);
			cs.setAlignment(HorizontalAlignment.CENTER);
			cs.setVerticalAlignment(VerticalAlignment.CENTER);
			cell.setCellType(CellType.STRING);
			cell.setCellStyle(cs);
			
			//Setting a cell to new k
			cell.setCellValue(bs.toString());
			
			//Creating export location
			String x = "fgh";
			File f2 = new File("C:\\Users\\samuel\\Documents\\Tupper's Formula\\"+ x +".xlsx");
			
			//Putting changes on export location
			FileOutputStream fileOut = new FileOutputStream(f2);
			wb.write(fileOut);
			fileOut.close();

			// Closing the workbook
			wb.close();
		}
		
		catch(IOException e) {
			//Just in case
			System.out.print(e);
		}
	}
}
