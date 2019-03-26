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
		String binaryString;
		File f;
		FileInputStream fis;
		Workbook wb;
		String b = "1";
		try{
			f = new File("C:\\Users\\samuel\\Documents\\Tupper's Formula\\Editable Formula2.xlsx");
			fis = new FileInputStream(f);
			wb = new XSSFWorkbook(fis);

			XSSFSheet sheet1 = (XSSFSheet) wb.getSheetAt(1);
			XSSFSheet sheet2 = (XSSFSheet) wb.getSheetAt(0);
			b = sheet1.getRow(5).getCell(2).getStringCellValue();
			
			binaryString = b.replace("	", "");
			BigInteger bs = new BigInteger(binaryString,2);
			bs = bs.multiply(new BigInteger("17"));
			
			XSSFCell cell = sheet2.createRow(25).createCell(2);
			XSSFCellStyle cs = (XSSFCellStyle) wb.getCellStyleAt(1);
			cs.setWrapText(true);
			cs.setAlignment(HorizontalAlignment.CENTER);
			cs.setVerticalAlignment(VerticalAlignment.CENTER);
			
			cell.setCellType(CellType.STRING);
			cell.setCellStyle(cs);
			cell.setCellValue(bs.toString());
			
			String x = "fgh";
			File f2 = new File("C:\\Users\\samuel\\Documents\\Tupper's Formula\\"+ x +".xlsx");
			
			FileOutputStream fileOut = new FileOutputStream(f2);
			wb.write(fileOut);
			fileOut.close();

			// Closing the workbook
			wb.close();
		}
		catch(IOException e) {
			System.out.print(e);
		}
	}
}