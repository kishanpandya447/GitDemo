import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDriven {

	

	public static void main(String[] args) throws IOException {
		
		XSSFWorkbook workbook=new XSSFWorkbook("D:\\exceldriven\\demodata.xlsx");
		
		int sheets= workbook.getNumberOfSheets();
		ArrayList<String> al= new ArrayList<String>();
		for(int i=0;i<sheets;i++) {
			if(workbook.getSheetName(i).equals("TestData")) {
				
				XSSFSheet sheet=workbook.getSheetAt(i);
				Iterator<Row> rows=sheet.rowIterator();
				Row firstrow = rows.next();
				Iterator<Cell> cells=firstrow.cellIterator();
				int k=0;
				int column=0;
				
				while(cells.hasNext()) {
				Cell value=cells.next();
				Cell value=cells.next();
				
			if(	value.getStringCellValue().equalsIgnoreCase("TestCases")) {
				
				System.out.println("Index of column is "+k);
				column=k;
				
			}
			k++;
				}
				
				while(rows.hasNext()) {
					
					Row r =rows.next();
					//System.out.println(r.getCell(column).getStringCellValue());
					if(r.getCell(column).getStringCellValue().equalsIgnoreCase("Purchase")) {
						Iterator<Cell> cv=r.cellIterator();
					while(cv.hasNext()) {
						
						Cell c=cv.next();
						if (c.getCellType()==CellType.STRING) {
						
					al.add(c.getStringCellValue());
						}
						else if(c.getCellType()==CellType.NUMERIC) {
							al.add(NumberToTextConverter.toText(c.getNumericCellValue()));
							
						}
						
					}
					
						
					}
				}
				
			}
			

		}
		
		Iterator<String> values=al.iterator();
		while(values.hasNext()) {
			System.out.println(values.next());
			
		}
	}

}
