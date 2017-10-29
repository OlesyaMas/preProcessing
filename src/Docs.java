import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Docs {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
			FileInputStream fis = new FileInputStream (new File("patient_los.xlsx"));
			
			//create workbook instance
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			
			//create a sheet object to retrieve the sheet
			XSSFSheet sheet = wb.getSheetAt(0);
			
			//for evaluating cell type
			FormulaEvaluator formulaEvaluator = wb.getCreationHelper().createFormulaEvaluator();
			for(Row row : sheet) {
				for(Cell cell : row) {
					switch(formulaEvaluator.evaluateInCell(cell).getCellType()) {
					//if cell is a numeric format
					case Cell.CELL_TYPE_NUMERIC:
						System.out.print(cell.getNumericCellValue() + "\t\t");
						break;
					case Cell.CELL_TYPE_STRING:
						System.out.print(cell.getStringCellValue() + "\t\t");
						break;
					}
				}
				System.out.println();
			}
			

	}
}