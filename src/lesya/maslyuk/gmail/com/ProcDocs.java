package lesya.maslyuk.gmail.com;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.Format;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import lesya.maslyuk.gmail.com.model.ColumnAttr;
import lesya.maslyuk.gmail.com.model.ColumnType;
import lesya.maslyuk.gmail.com.utils.MyUtils;

public class ProcDocs {
	
	private static final int FIRST_ROW = 1;
	private static int lastRowNum = 0;
	private static final String COLUMN_CODE = "_code";
	private static final int TITLE_ROW_NUM = 0;	//in Excel #1
	private static final int SHIFT_FORMULA_ROW = 2;
	private static final int PROCEED_ROWS = 10;
	public static boolean DELETE_EMPTY = false;
	public static boolean CALC_AVERAGE = false;
	
	///private static Map<Integer,Row> rowMap = new HashMap<Integer,Row>();
	private static DataFormatter objDefaultFormat = new DataFormatter();
	private static FormulaEvaluator objFormulaEvaluator;
	private static Map<String,Integer> titleMap = new HashMap<String,Integer>();
	private static Map<Integer,String> indexMap = new HashMap<Integer,String>();
	private static Map<String,ColumnAttr> attrMap = new LinkedHashMap<String,ColumnAttr>();

	private static XSSFWorkbook myWorkBook; 
	private static XSSFSheet mySheet;
	

	//String inputFile = "D:\\Dev\\neon_ws\\OlesyaPrj\\data\\InputDocs1.xls";
	///static String inputFile = "D:\\study\\java_proj\\DataProc\\data\\Книга1_int-1.xlsx";
	static String inputFile = "D:\\study\\java_proj\\DataProc\\data\\Книга1_1000.xlsx";
	static String outputFile = "D:\\study\\java_proj\\DataProc\\data\\Книга1_out_test.xlsx";
	
/*	public static void main(String[] args) {
		// TODO Auto-generated method stub
		System.out.println("ProcDocs.main()");
		try {
			
			myWorkBook = new XSSFWorkbook (new FileInputStream(inputFile)); // Return first sheet from the XLSX workbook
			mySheet = myWorkBook.getSheetAt(0);
			
			lastRowNum = mySheet.getLastRowNum();
			System.out.println("lastRowNum=" + lastRowNum);
			
			readFromExcel(inputFile);
			//System.out.println("rowMap.size()=" + rowMap.size());
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
*/
	
	public static void readFromExcel(String inputFile, String outputFile) throws Exception{
		ProcDocs.inputFile = inputFile;
		ProcDocs.outputFile = outputFile;
		
		myWorkBook = new XSSFWorkbook (new FileInputStream(inputFile)); // Return first sheet from the XLSX workbook
		mySheet = myWorkBook.getSheetAt(0);
		
		lastRowNum = mySheet.getLastRowNum();
		System.out.println("lastRowNum=" + lastRowNum);
		
		
//        HSSFWorkbook myExcelBook = new HSSFWorkbook(new FileInputStream(file));
//        HSSFSheet myExcelSheet = myExcelBook.getSheet("Stat");
//        HSSFRow firstRow = myExcelSheet.getRow(0);
		XSSFRow firstRow = mySheet.getRow(0);
		
        String title = firstRow.getCell(0).getStringCellValue();        
        System.out.println(title);
    
        //FormulaEvaluator objFormulaEvaluator = new HSSFFormulaEvaluator((HSSFWorkbook) myWorkBook);
        //FormulaEvaluator objFormulaEvaluator = new XSSFFormulaEvaluator(myWorkBook);
        objFormulaEvaluator = new XSSFFormulaEvaluator(myWorkBook);
        
        Row titleRow = mySheet.getRow(TITLE_ROW_NUM);
        System.out.println("title Row " + titleRow.getRowNum());
		defineTitleMap(titleRow);
	    System.out.println(" ");
	    
//D        
//	    Set<String> columnsName = titleMap.keySet();
//	    Iterator it = columnsName.iterator();
//	    while(it.hasNext()){
//	    	System.out.println(it.next());
//	    }
	    
        for (Row row : mySheet) {
        	int zeroLine = row.getRowNum();
        	System.out.println("zero Row " + zeroLine);
        	if (zeroLine == 0){
        		continue;
        	}
        	
        	//TODO fix First 1000
        	//if (firstLine == PROCEED_ROWS)
        	//	break;

        	//System.out.print("row.getRowNum()=" + row.getRowNum() + "||");
        	//rowMap.put(row.getRowNum(), row);
        	
        	//defineTypes
            defineColumnTypes(row);
            System.out.println();
        }

//DEBUG        
/*	    Set entrySet = attrMap.entrySet();
	    Iterator itEntries = entrySet.iterator();
	    while(itEntries.hasNext()){
	    	Map.Entry entry = (Map.Entry)itEntries.next();
	    	ColumnAttr value = (ColumnAttr)entry.getValue();
			System.out.println(entry.getKey() + " " + value.getColumnType());
			
			Map<ColumnType, Integer> typeMap = value.getTypeMap();
			Set entrySetT = typeMap.entrySet();
			Iterator itTypes = entrySetT.iterator();
			while(itTypes.hasNext()){
		    	Map.Entry mTypes = (Map.Entry)itTypes.next();
		    	System.out.println(mTypes.getKey() + "   " + mTypes.getValue());
			}
	    }
*/

        /*
         * Calculate average for DECIMAL
         */
        if(CALC_AVERAGE)
        	calcAverage(titleRow);

        int lastColumn = titleRow.getLastCellNum();
        System.out.println("DEBUG:0: row.getLastCellNum()=" + lastColumn);
        
        LinkedList<Integer> rowsToRemove = new LinkedList<Integer>();
        for (Row row : mySheet) {
        	int curRowIndex = row.getRowNum();
        	if (curRowIndex == 0){
        		continue;
        	}
//        	if (curRowIndex == 1){
//            	//Row oneRow = mySheet.getRow(FIRST_ROW);
//                CellStyle currentStyle = row.get .getCellStyle();        
//        	}
        	System.out.println("DEBUG:000000000000000000000000000000000000000000000): row.getRowNum()=" + row.getRowNum());
        	
        	//CREATE_NULL_AS_BLANK vs Row.RETURN_BLANK_AS_NULL
        	cycleCell: 
        	for (int ci = 0; ci < lastColumn; ci++) {
    			String columnName = indexMap.get(ci);
	    		System.out.println("DEBUG:1: " + columnName);

	    		if(columnName != null){	//skip columns with empty title
    		    	ColumnAttr columnAttr = attrMap.get(columnName);
    		    	ColumnType curColType = columnAttr.getColumnType();
    		    	System.out.println("DEBUG:2: curColType =" + curColType);
    		    	
	                Cell cell = row.getCell(ci, Row.CREATE_NULL_AS_BLANK);
	                CellStyle currentStyle = columnAttr.getCellStyle(); 
	                if (MyUtils.isCellEmpty(cell)) {
	                   // The spreadsheet is empty in this cell
    		    		System.out.println("DEBUG:3: Cell is empty !!!!!!!!!!!!!!!!!!!!!!!");

    		    		if(CALC_AVERAGE){
        		    		if(curColType == ColumnType.DECIMAL || (curColType == ColumnType.INTEGER && !columnAttr.isColumnOrdinal())){
            		    		System.out.println("DEBUG:44: = set CELL VALUE: columnAttr.getAverage()" + columnAttr.getAverage());
            		    		
            		    		//Cell cellNew = row.createCell(ci);
            		    		//cellNew.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
            		    		//cellNew.setCellValue(columnAttr.getAverage());
            		    		
            		    		//Cell cell = row.getCell(2);
            		    	    //if (cell == null) {
            		    	        cell = row.createCell(ci);
            		    	    //}
            		    	    cell.setCellType(Cell.CELL_TYPE_NUMERIC);
            		    	    cell.setCellStyle(currentStyle);
            		    	    cell.getCellStyle().setWrapText(false);
            		    	    cell.setCellValue(columnAttr.getAverage());
            		    	}
    		    		}

    		    		if(DELETE_EMPTY){
        		    		System.out.println("DEBUG:5: Cell will be deleted !!! " + curRowIndex);
        		    		if(!rowsToRemove.contains(curRowIndex))
        		    			rowsToRemove.addFirst(curRowIndex);
        		    		
        		    		//TODO use this to performence (if empty cell => the other cells stay empty !!!)
        		            //if(DELETE_EMPTY){
        		            //	break cycleCell;
        		            //}
        		    	}
    		    		//TODO other fixes
    		    		
	                }
	    		}
                    // Do something useful with the cell's contents
                	//cell.setCellStyle(currentStyle);
		    	    //cell.getCellStyle().setWrapText(false);
             }
        }
//        if(DELETE_EMPTY){
//    		for (Integer rowInd : rowsToRemove) {
//    			removeRow(mySheet, rowInd);
//    		}
//        }
        
        saveNew();
        	
    	myWorkBook.close();
    }


	private static void calcAverage(Row titleRow) throws Exception {
		//Row titleRow = mySheet.getRow(0);
		//Row oneRow = mySheet.getRow(FIRST_ROW);
		//Row lastRow = mySheet.getRow(lastRowNum);
		XSSFRow rowResult = mySheet.createRow(lastRowNum + SHIFT_FORMULA_ROW);
		
		for (Cell cell : titleRow){
			int curColumnIndex = cell.getColumnIndex();
			String columnName = indexMap.get(curColumnIndex);            	
			if(columnName != null){	//skip columns with empty title
				ColumnAttr columnAttr = attrMap.get(columnName);
		    	if(columnAttr.getColumnType() == ColumnType.DECIMAL || (columnAttr.getColumnType() == ColumnType.INTEGER && !columnAttr.isColumnOrdinal())){
		    		System.out.println("------------- for column " + columnName);
		    		
		    		//Cell cellBeginRange = oneRow.getCell(curColumnIndex);
		    		
		    		Double avg = null;
		    		//if(columnAttr.getTypeMap().get(ColumnType.DECIMAL) > 1){
		        		//Cell endRange = lastRow.getCell(curColumnIndex);
		    
		        		//if(cellBeginRange == null)
		        		//	throw new Exception("DEBUG: cellBeginRange is null");

		        		//if(endRange == null)
		        			//throw new Exception("DEBUG: endRange is null");
		        		
		        		avg = avg(mySheet, myWorkBook,curColumnIndex, rowResult);

//                		}else{
//                    		avg = avg(mySheet, myWorkBook, cellBeginRange, cellBeginRange);
		    		//}
		        		
		        	if(columnAttr.getColumnType() == ColumnType.INTEGER){
		        		int intAvg = (int)Math.round(avg);
		        		columnAttr.setAverage(intAvg);
		        	}else{
		        		columnAttr.setAverage(avg);
		        	}
		        	
		    		                		
		    		System.out.println("AVERAGE for column " + columnName + " : " + columnAttr.getAverage());
		    	}
		    }
		}
		mySheet.removeRow(rowResult);
	}


	private static void defineColumnTypes(Row row) throws Exception {
		for (Cell cell : row) {                	
			String columnName = indexMap.get(cell.getColumnIndex());
			if(columnName != null){	//skip columns with empty title
		    	ColumnAttr columnAttr = attrMap.get(columnName);
		    	calcCellType(cell, columnAttr);
		    	if(!MyUtils.isCellEmpty(cell))
		    		columnAttr.setCellStyle(cell.getCellStyle());
		    	
		    	System.out.print("|");
		    	
//                    	DataFormatter dataFormatter = new DataFormatter();
//                    	Format cellFDefault = dataFormatter.getDefaultFormat(cell); 
//                    	System.out.println(cellFDefault.toString());
			}
		}
	}


	private static void defineTitleMap(Row row) {
		for (Cell cell : row){
			objFormulaEvaluator.evaluate(cell); // This will evaluate the cell, And any type of cell will return string value
		    String cellValueStr = objDefaultFormat.formatCellValue(cell,objFormulaEvaluator);
		    System.out.print(cellValueStr + " " + cell.getColumnIndex() + "|");
			if(cell.getCellType() == Cell.CELL_TYPE_BLANK || cellValueStr.contains(COLUMN_CODE))
				continue;
			
			//String colName = CellReference.convertNumToColString(cell.getColumnIndex());
			//--System.out.print(cell.getStringCellValue() + " " + cell.getColumnIndex() + "|");
		    titleMap.put(cellValueStr, cell.getColumnIndex());
		    indexMap.put(cell.getColumnIndex(), cellValueStr);
		    
		    attrMap.put(cellValueStr, new ColumnAttr(cell.getColumnIndex()));
		}
	}

/*
	private static void procRow(Cell cell) {
		if(cell.getCellType() == HSSFCell.CELL_TYPE_STRING){
		    String cellValue = cell.getStringCellValue();
		    //System.out.print("cellValue : " + cellValue);
		    System.out.print(cellValue + " ");
		}
		
//		if(cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC){
//			Date birthdate = cell.getDateCellValue();
//			System.out.print("birthdate :" + birthdate);
//		}
	}
*/	
	
	
	private static void calcCellType(Cell cell, ColumnAttr columnAttr) throws Exception {
		//System.out.println("DEBUG:" + indexMap.get(columnAttr.getIndex()));
		//Cell cell = row.getCell(ci, Row.CREATE_NULL_AS_BLANK);
		if(MyUtils.isCellEmpty(cell)){
			System.out.print("Empty: ");
			incAttrType(columnAttr, ColumnType.BLANK);
			return;
		}
		
		switch (cell.getCellType()) {
			case Cell.CELL_TYPE_BLANK:
				System.out.print("BLANK: ");
				incAttrType(columnAttr, ColumnType.BLANK);
				break;		
			case Cell.CELL_TYPE_STRING:
				//failuresCell.setCellValue(cell.getRichStringCellValue().getString());
				System.out.print("T: " + cell.getRichStringCellValue().getString());
				incAttrType(columnAttr, ColumnType.TEXT);
				break;
			case Cell.CELL_TYPE_NUMERIC:
				if (DateUtil.isCellDateFormatted(cell)) {
					//failuresCell.setCellValue(cell.getDateCellValue());
					System.out.print("D: " + cell.getDateCellValue());
					incAttrType(columnAttr, ColumnType.DATE);
				} else {
					//failuresCell.setCellValue(cell.getNumericCellValue());
					Double numericCellValue = cell.getNumericCellValue();
					if(Math.floor(numericCellValue) == numericCellValue){
						System.out.print("I: " + Math.floor(numericCellValue));
						incAttrType(columnAttr, ColumnType.INTEGER);
					}else{
						System.out.print("N: " + cell.getNumericCellValue());
						incAttrType(columnAttr, ColumnType.DECIMAL);
					}       
				}
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				//failuresCell.setCellValue(cell.getBooleanCellValue());
				System.out.print("B: " + cell.getBooleanCellValue());
				incAttrType(columnAttr, ColumnType.BOOLEAN);
				break;
			case Cell.CELL_TYPE_FORMULA:
				//TODO incAttrType(columnAttr, ColumnType );
				//failuresCell.setCellValuec(cell.getCellFormula());
				System.out.print("FX: " + cell.getCellFormula());
				break;
			default:
         // some code
		}
		
/*		
		//System.out.print("[" + objDefaultFormat.formatCellValue(cell) + "]");
		String cellValueStr = objDefaultFormat.formatCellValue(cell,objFormulaEvaluator);
		System.out.print("[[" + cellValueStr + "]]");
		
		//analysis
		Exception exp = new Exception("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
		int curColInd = cell.getColumnIndex();
		if(curColInd == titleMap.get("Num").intValue()){
			if(cell.getCellType() == Cell.CELL_TYPE_STRING){
				System.out.println("!!!!!!!!!!!!!!!!!!!!!! Num");
				//throw exp;
			}
		}
		if(curColInd == titleMap.get("ID").intValue()){
			if(cell.getCellType() == Cell.CELL_TYPE_STRING){
				System.out.println("!!!!!!!!!!!!!!!!!!!!!! ID");
				//throw exp;
			}
		}
		if(curColInd == titleMap.get("PPT").intValue()){
			if(cell.getCellType() == Cell.CELL_TYPE_STRING){
				System.out.println("!!!!!!!!!!!!!!!!!!!!!! PPT");
				//throw exp;
			}
		}
		if(curColInd == titleMap.get("IMT").intValue()){
			if(cell.getCellType() == Cell.CELL_TYPE_STRING){
				System.out.println("!!!!!!!!!!!!!!!!!!!!!! IMT");
				//throw exp;
			}
		}
*/
		
//		if(curColInd == titleMap.get("KDI").intValue()){
//			if(cell.getCellType() == Cell.CELL_TYPE_STRING || cell.getCellType() == Cell.CELL_TYPE_BLANK){
//				System.out.println("!!!!!!!!!!!!!!!!!!!!!! KDI");
//				throw exp;
//			}
//		}
		

	}

	private static void incAttrType(ColumnAttr columnAttr, ColumnType columnType) {
		//System.out.println("DEBUG:" + indexMap.get(columnAttr.getIndex()));
		if(columnAttr.getTypeMap().containsKey(columnType)){
			int curValue = columnAttr.getTypeMap().get(columnType).intValue() + 1;
			columnAttr.getTypeMap().put(columnType, curValue);
		}else{
			columnAttr.getTypeMap().put(columnType, 1);
		}
	}

//	private static String getCellName(Cell cell){
//		System.out.println("DEBUG 1 cell.getColumnIndex()=" + cell.getColumnIndex());
//		System.out.println("DEBUG 2 cell.getRowIndex()=" + cell.getRowIndex());
//	    return CellReference.convertNumToColString(cell.getColumnIndex()) + (cell.getRowIndex() + 1);
//	}

	private static String getCellName(int curColumnIndex, int rowIndex){
	    return CellReference.convertNumToColString(curColumnIndex) + (rowIndex + 1);
	}

	
	private static double avg(XSSFSheet mySheet, XSSFWorkbook myWorkBook, int curColumnIndex, XSSFRow rowResult){
		//XSSFRow rowResult = mySheet.createRow(lastRowNum + SHIFT_FORMULA_ROW);
		//Cell cell = rowResult.createCell(1);
		Cell cell = rowResult.createCell(curColumnIndex);
		//System.out.println("getCellName(beginRange)=" + getCellName(beginRange));
		//System.out.println("getCellName(endRange)=" + getCellName(endRange));
		String range = getCellName(curColumnIndex,FIRST_ROW) + ":" + getCellName(curColumnIndex,lastRowNum);	//C2:C10
		System.out.println("Range=" + range);
		
		cell.setCellType(XSSFCell.CELL_TYPE_FORMULA);
		cell.setCellFormula("AVERAGE(" + range + ")");	    
		//cell = row.createCell(3);
		//cell.setCellValue("SUM(O2:O10)");
		XSSFFormulaEvaluator evaluator = myWorkBook.getCreationHelper().createFormulaEvaluator();
		XSSFFormulaEvaluator.evaluateAllFormulaCells(myWorkBook);
		CellValue cellValue = evaluator.evaluate(cell);
		System.out.println("FUNC prepared = " + cellValue.formatAsString());
		//String result = cellValue.formatAsString(); //cellValue.getNumberValue()
		double result = cellValue.getNumberValue();
		//mySheet.removeRow(rowResult);
		
		return result;
	}

	
	public static void removeRow(Sheet sheet, int rowIndex) {
	    int lastRowN = sheet.getLastRowNum();
	    if (rowIndex > 0 && rowIndex < lastRowN) {
	        sheet.shiftRows(rowIndex + 1, lastRowN, -1);
	    }
	    if (rowIndex == lastRowN) {
	        Row removingRow = sheet.getRow(rowIndex);
	        if (removingRow != null) {
	            sheet.removeRow(removingRow);
	        }
	    }
	}
	
	private static void saveNew(){
		File outWB = new File(outputFile);
		OutputStream out = null;
		try {
			out = new FileOutputStream(outWB);
			myWorkBook.write(out);
						
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}finally{
			try {
				out.flush();
				out.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}

	}
	
	
	public void applyStyleToRange(Sheet sheet, CellStyle style, int rowStart, int colStart, int rowEnd, int colEnd) {
	    for (int r = rowStart; r <= rowEnd; r++) {
	        for (int c = colStart; c <= colEnd; c++) {
	            Row row = sheet.getRow(r);

	            if (row != null) {
	                Cell cell = row.getCell(c);

	                if (cell != null) {
	                    cell.setCellStyle(style);
	                }
	            }
	        }
	    }
	}
}
	
//TODO
//fix processing columns without title
//calc average
//remove empty	
//save file	
//Vozrast - remove 4792 + Blanks
//KDO1 - max 177301
//FV1 - exclude extra (1 DECIMAL)




//String columnName = indexMap.get(cell.getColumnIndex());
//ColumnAttr columnAttr = attrMap.get(columnName);

//System.out.println("column Name=" + columnName);
//System.out.println("columnAttr.getColumnType()=" + columnAttr.getColumnType());

//String columnName = indexMap.get(cell.getColumnIndex());
//ColumnAttr columnAttr = attrMap.get(columnName);
//
//if(cell.getCellType() == Cell.CELL_TYPE_BLANK)
//continue;


//}


/*        for (Iterator<Row> iterator = myExcelSheet.iterator(); iterator.hasNext();) {
HSSFRow row = (HSSFRow) iterator.next();

if(row.getCell(0).getCellType() == HSSFCell.CELL_TYPE_STRING){
    String name = row.getCell(0).getStringCellValue();
    System.out.println("name : " + name);
}
}
*/        

//if(row.getCell(1).getCellType() == HSSFCell.CELL_TYPE_NUMERIC){
//Date birthdate = row.getCell(1).getDateCellValue();
//System.out.println("birthdate :" + birthdate);
//}






/*
cycleCell: 
for (Cell cell : row) {    			
	String columnName = indexMap.get(cell.getColumnIndex());
	System.out.println("DEBUG:1: " + columnName);
	
	if(columnName != null){	//skip columns with empty title
    	ColumnAttr columnAttr = attrMap.get(columnName);
    	ColumnType curColType = columnAttr.getColumnType();
    	if(MyUtils.isCellEmpty(cell)){
    		System.out.println("DEBUG:2: CELL_TYPE_BLANK ");
    		System.out.println("DEBUG:3: curColType =" + curColType);
    		
    		if(curColType == ColumnType.DECIMAL || (curColType == ColumnType.INTEGER && !columnAttr.isColumnOrdinal())){
	    		System.out.println("DEBUG:4: = value set  ");
	    		cell.setCellValue(columnAttr.getAverage());
	    	}else{
	    		if(!rowsToRemove.contains(curRowIndex))
	    			rowsToRemove.add(curRowIndex);
	    		break cycleCell;
	    	}
	    	//TODO other fixes
    	}
	}
}

*/