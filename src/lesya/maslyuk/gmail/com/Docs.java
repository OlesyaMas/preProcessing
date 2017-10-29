package lesya.maslyuk.gmail.com;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Docs {

	public static void main(String[] args) throws IOException {
		//readFile();
			

	}

	public static void readFile(File inputFile) throws FileNotFoundException, IOException, InvalidFormatException {
		// TODO Auto-generated method stub
			//FileInputStream fis = new FileInputStream (new File("data/test.xlsx"));
			
			//create workbook instance
			XSSFWorkbook wb = new XSSFWorkbook(inputFile);
			
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