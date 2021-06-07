import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDriven2 {

	

	public static void main(String[] args) throws IOException {
		
		XSSFWorkbook workbook= new XSSFWorkbook("D:\\exceldriven\\demodata.xlsx");
		int totalSheets=workbook.getNumberOfSheets();
		int columncount=0;
		for (int i=0;i<totalSheets;i++) {
			
			
			if(workbook.getSheetName(i).equals("TestData")) {
				XSSFSheet singlesheet =workbook.getSheetAt(i);
			Iterator<Row> rows=singlesheet.rowIterator();
			Row singleFirstRow=	rows.next();
			Iterator<Cell> cells=singleFirstRow.cellIterator();
			while(cells.hasNext()) {
			Cell singleCell=cells.next();
			
			if (singleCell.getStringCellValue().equals("data3")) {
				
				System.out.println("found data3 at position "+columncount);
				break;
			}
			columncount=columncount+1;
			}
			
			while(rows.hasNext()) {
				
				Row singleRow=rows.next();
				System.out.println(singleRow.getCell(columncount).getStringCellValue());				
			}
			
			
			break;
			}
			
		}
		
	}


		
	}