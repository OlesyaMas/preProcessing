package lesya.maslyuk.gmail.com;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.Format;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.NavigableSet;
import java.util.Set;
import java.util.SortedSet;
import java.util.TreeSet;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import lesya.maslyuk.gmail.com.TestClass.ShiftCells;
import lesya.maslyuk.gmail.com.model.ColumnAttr;
import lesya.maslyuk.gmail.com.model.ColumnType;
import lesya.maslyuk.gmail.com.model.MaxValue;
import lesya.maslyuk.gmail.com.model.MinValue;
import lesya.maslyuk.gmail.com.utils.MyUtils;

public class Processing {
	
	private static final int MIN_MAX_LOWER = 5;
	private static final int FIRST_ROW = 1;
	//private static int lastRowNum = 0;
	private static final String COLUMN_CODE = "_code";
	private static final int TITLE_ROW_NUM = 0;	//in Excel #1
	private static final int SHIFT_FORMULA_ROW = 10;
	private static final int PROCEED_ROWS = 10;
	public static boolean DELETE_EMPTY = false;
	public static boolean CALC_AVERAGE = false;
	public static boolean REMOVE_MIN_MAX = false;
	public static boolean CODING = false;
	
	private static DataFormatter objDefaultFormat = new DataFormatter();
	private static FormulaEvaluator objFormulaEvaluator;
	
	//TODO + 1 for maps: titleMap,indexMap,   attrMap?
	private static Map<String,Integer> titleMap = new HashMap<String,Integer>();
	private static Map<Integer,String> indexMap = new HashMap<Integer,String>();
	private static Map<String,ColumnAttr> attrMap = new LinkedHashMap<String,ColumnAttr>();

	private static XSSFWorkbook myWorkBook; 
	private static XSSFSheet mySheet;
	
	
	///static String inputFile = "D:\\study\\java_proj\\multi\\ ÌË„‡1_100_nocodes.xlsx";
	///static String outputFile = "D:\\study\\java_proj\\multi\\ ÌË„‡1_out_nocodes_test.xlsx";
	static SortedSet<Integer> rowsToRemove = new TreeSet<Integer>();
	
	static String inputFile;
	static String outputFile;
	
	
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		System.out.println("ProcDocs.main()");
		try {
			
			myWorkBook = new XSSFWorkbook (new FileInputStream(inputFile)); // Return first sheet from the XLSX workbook
			mySheet = myWorkBook.getSheetAt(0);
			
			System.out.println("lastRowNum=" + mySheet.getLastRowNum());
			
			//TO USE
			//stayleOrigin.getDataFormat()!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
			
			//removeRow(mySheet, mySheet.getLastRowNum());
			//removeRow(mySheet, mySheet.getLastRowNum() - 10);
/*			for (Row row : mySheet) {
				for (Cell cell : row){
					//cell.setCellStyle(stayleOrigin);
				}
			}
*/			
			readFromExcel(inputFile, outputFile);
			//System.out.println("rowMap.size()=" + rowMap.size());
			
			//saveNew();
			
			//myWorkBook.close();
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	
	
	public static void readFromExcel(String inputFile, String outputFile) throws Exception{
		Processing.inputFile = inputFile;
		Processing.outputFile = outputFile;
		
		myWorkBook = new XSSFWorkbook (new FileInputStream(inputFile)); // Return first sheet from the XLSX workbook
		mySheet = myWorkBook.getSheetAt(0);
		
		System.out.println("lastRowNum=" + mySheet.getLastRowNum());
		
		
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
	    
//        
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
        
        if(REMOVE_MIN_MAX){
        	System.out.println("DEBUG 110 processing min/max");
	        for (Row row : mySheet) {
	        	int zeroLine = row.getRowNum();
	        	if (zeroLine == 0)
	        		continue;
	        	
	        	for (Cell cell : row) {                	
	    			String columnName = indexMap.get(cell.getColumnIndex());
	    			if(columnName != null){	//skip columns with empty title
	    		    	ColumnAttr columnAttr = attrMap.get(columnName);
			    		if(!MyUtils.isCellEmpty(cell) && columnAttr.isColNumCalculatable()){
					    	fillMinMax(row, cell, columnAttr);		    			
			    		}
	    			}
	    		}            
	        }
	        
			for (Cell cell : titleRow){
				int curColumnIndex = cell.getColumnIndex();
				String columnName = indexMap.get(curColumnIndex);            	
				if(columnName != null){	//skip columns with empty title
					ColumnAttr columnAttr = attrMap.get(columnName);
					System.out.println("DEBUG XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX " + columnName);

			    	if(columnAttr.isMinMaxRemovingApplicable()){
			    		rowsToRemove.add(columnAttr.getMinValue().getRowNum());
			    		rowsToRemove.add(columnAttr.getMaxValue().getRowNum());
			    		System.out.println("DEBUG 112 MIN:" + columnAttr.getMinValue().toString());
			    		System.out.println("DEBUG 113 MAX:" + columnAttr.getMaxValue().toString());
			    	}
			    }
			}
	        removeEmptyRows();
        }

			    	
        /*
         * for DECIMAL
         */
        if(CALC_AVERAGE)
        	calcAverage(titleRow);

        int lastColumn = titleRow.getLastCellNum();
        System.out.println("DEBUG:0: row.getLastCellNum()=" + lastColumn);
        
        for (Row row : mySheet) {
        	int curRowIndex = row.getRowNum();
        	if (curRowIndex == 0) continue;
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
    		    	
	                Cell cell = row.getCell(ci, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK); 
	                CellStyle currentStyle = columnAttr.getCellStyle(); 
	                if (MyUtils.isCellEmpty(cell)) {
    		    		System.out.println("DEBUG:3: Cell is empty !!!!!!!!!!!!!!!!!!!!!!!");

    		    		if(CALC_AVERAGE){
        		    		if(columnAttr.isColNumCalculatable()){
            		    		System.out.println("DEBUG:44: = set CELL VALUE: columnAttr.getAverage()" + columnAttr.getAverage());
            		    		
            		    	    cell = row.createCell(ci);
            		    	    cell.setCellType(Cell.CELL_TYPE_NUMERIC);
            		    	    cell.setCellStyle(currentStyle);
            		    	    cell.getCellStyle().setWrapText(false);
            		    	    cell.setCellValue(columnAttr.getAverage());
            		    	}
    		    		}

    		    		if(DELETE_EMPTY){
        		    		System.out.println("DEBUG:5: Cell will be deleted !!! " + curRowIndex);
        		    		if(!rowsToRemove.contains(curRowIndex))
        		    			rowsToRemove.add(curRowIndex);
        		    		
        		    		//TODO use this to performence (if empty cell => the other cells stay empty !!!)
        		            //	break cycleCell;
        		    	}
	                }
	    		}
                //Do something useful with the cell's contents
             }
        }

        if(DELETE_EMPTY){
        	removeEmptyRows();
        }
        
        if(CODING){
        	CodingProcessor codingProcessor = new CodingProcessor();
        	codingProcessor.process(mySheet, indexMap, attrMap);
        	
			//loop for title columns
        	Set<String> titles = titleMap.keySet();
        	for (String columnTitle : titles) {
        		int curColumnIndex = getTitleColumnIndex(columnTitle);
				ColumnAttr columnAttr = attrMap.get(columnTitle);
				System.out.println("DEBUG CODING columnName = " + columnTitle);
				
		    	if(!columnAttr.isColumnCodeable()){
		    		continue;
		    	}
				
				int nrRows = mySheet.getLastRowNum()+1;
				int nrCols = mySheet.getRow(0).getLastCellNum();
				
				//shift all rows for defined Coding Columns
				for (int row = 0; row < nrRows; row++) {
					Row r = mySheet.getRow(row);

					if (r == null) {
						continue;
					}
					
					//shift right cells
					TestClass testClass = new TestClass();
					int rightColIndex = curColumnIndex + 1;
					//TestClass.shifColumn(nrCols, rightColIndex, r);
					testClass.shifColumn(nrCols, rightColIndex, r);
										
					//create new column
					Cell curr = r.getCell(curColumnIndex);					
					int codeValue = -1;
					if (r.getRowNum() == TITLE_ROW_NUM) {
						objFormulaEvaluator.evaluate(curr); // This will evaluate the cell, And any type of cell will return string value
					    String cellValueStr = objDefaultFormat.formatCellValue(curr,objFormulaEvaluator) + COLUMN_CODE;
						int cellType = curr.getCellType();
						r.createCell(rightColIndex, cellType);
						r.getCell(rightColIndex).setCellValue(cellValueStr);
						r.getCell(rightColIndex).setCellStyle(curr.getCellStyle());
					}else{
						codeValue = columnAttr.getCodingMap().getCode(curr.getStringCellValue());
						if (codeValue >= 0){
							int cellType = Cell.CELL_TYPE_NUMERIC;
							r.createCell(rightColIndex, cellType);
							r.getCell(rightColIndex).setCellValue(codeValue);
						}else{
							int cellType = Cell.CELL_TYPE_BLANK;
							r.createCell(rightColIndex, cellType);
						}
					}
					
				}
				
			}
        }
        
        
        
        saveNew();
        	
    	myWorkBook.close();
    }



	private static Integer getTitleColumnIndex(String title) {
		Row dynTitleRow = mySheet.getRow(TITLE_ROW_NUM);
		Integer titleColumnIndex = null;
		for (Cell cell : dynTitleRow) {
			objFormulaEvaluator.evaluate(cell); // This will evaluate the cell, And any type of cell will return string value
		    String cellValueStr = objDefaultFormat.formatCellValue(cell,objFormulaEvaluator);
		    if(cellValueStr.equals(title)){
		    	titleColumnIndex = cell.getColumnIndex();
		    	break;
		    }else{
		    	continue;
		    }
		}
		return titleColumnIndex;
	}


	private static void removeEmptyRows() {
		NavigableSet<Integer> rowsToRemove_nav = ((TreeSet)rowsToRemove).descendingSet();	//because shifting from last rows
		for (Integer rowInd : rowsToRemove_nav) {
			System.out.println("DEBUG: 101 : will be removed rowInd =" + rowInd);
			removeRow(mySheet, rowInd);
		}
		
		for (Row row : mySheet) {
			for (Cell cell : row){
	        	int curRowIndex = row.getRowNum();
	        	if (curRowIndex == 0) continue;

	        	int curColumnIndex = cell.getColumnIndex();
				String columnName = indexMap.get(curColumnIndex);            	
				if(columnName != null){	//skip columns with empty title
					ColumnAttr columnAttr = attrMap.get(columnName);
					cell.setCellStyle(columnAttr.getCellStyle());
					cell.getCellStyle().setWrapText(false);
				}
			}
		}
	}
	
	
	private static void calcAverage(Row titleRow) throws Exception {
		//Row titleRow = mySheet.getRow(0);
		//Row oneRow = mySheet.getRow(FIRST_ROW);
		//Row lastRow = mySheet.getRow(lastRowNum);
		int lastRowNum = mySheet.getLastRowNum();
		XSSFRow rowResult = mySheet.createRow(lastRowNum + SHIFT_FORMULA_ROW);
		System.out.println("------------- XSSFRow rowResult.getRowNum() " + rowResult.getRowNum());
		XSSFFormulaEvaluator evaluator = myWorkBook.getCreationHelper().createFormulaEvaluator();
		evaluator.clearAllCachedResultValues();
		
		for (Cell cell : titleRow){
			int curColumnIndex = cell.getColumnIndex();
			String columnName = indexMap.get(curColumnIndex);            	
			if(columnName != null){	//skip columns with empty title
				ColumnAttr columnAttr = attrMap.get(columnName);
		    	if(columnAttr.isColNumCalculatable()){
		    		System.out.println("------------- for column " + columnName);
		    		Double avg = averageByColumn(mySheet, myWorkBook,curColumnIndex, lastRowNum, rowResult, evaluator);
		        		
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


//	private static boolean isColumnCalculatable(ColumnAttr columnAttr) {
//		return columnAttr.getColumnType() == ColumnType.DECIMAL || (columnAttr.getColumnType() == ColumnType.INTEGER && !columnAttr.isColumnOrdinal());
//	}


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


	private static void fillMinMax(Row row, Cell cell, ColumnAttr columnAttr) {
		
		if(columnAttr.isColNumCalculatable()){
			System.out.println("DEBUG 103: " + cell.getCellType() + " #############");
			Double numericCellValue = cell.getNumericCellValue();
			int rowNum = row.getRowNum();
			if(columnAttr.getMinValue() == null && columnAttr.getMaxValue() == null){
				System.out.print("DEBUG 104: ");
				 columnAttr.setMinValue(new MinValue(rowNum, numericCellValue));
				 columnAttr.setMaxValue(new MaxValue(rowNum, numericCellValue));
			 }else if(numericCellValue < columnAttr.getMinValue().getValue()){
				 System.out.print("DEBUG 105: ");
				 columnAttr.getMinValue().setValue(numericCellValue);
				 columnAttr.getMinValue().setRowNum(rowNum);
			 }else if(numericCellValue > columnAttr.getMaxValue().getValue()){
				 System.out.print("DEBUG 106: ");
				 columnAttr.getMaxValue().setValue(numericCellValue);
				 columnAttr.getMaxValue().setRowNum(rowNum);
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


	private static double averageByColumn(XSSFSheet mySheet, XSSFWorkbook myWorkBook, int curColumnIndex, int lastRowNum, XSSFRow rowResult, XSSFFormulaEvaluator evaluator){
		Cell cell = rowResult.createCell(curColumnIndex);
		String range = getCellName(curColumnIndex,FIRST_ROW) + ":" + getCellName(curColumnIndex, lastRowNum);	//C2:C10
		System.out.println("Range=" + range);
		
		cell.setCellType(XSSFCell.CELL_TYPE_FORMULA);
		cell.setCellFormula("AVERAGE(" + range + ")");	    
		CellValue cellValue = evaluator.evaluate(cell);
		System.out.println("FUNC prepared = " + cellValue.formatAsString());
		//String result = cellValue.formatAsString(); //cellValue.getNumberValue()
		return cellValue.getNumberValue();
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


/*
//=Ã»Õ(H2:H200),=—◊®“(H2:H200)
private static void calcMinMax(Row titleRow) throws Exception {
	SortedSet<Integer> setRemoveMinMax = new TreeSet<Integer>();
	XSSFRow rowResult = mySheet.createRow(lastRowNum + SHIFT_FORMULA_ROW);
	for (Cell cell : titleRow){
		int curColumnIndex = cell.getColumnIndex();
		String columnName = indexMap.get(curColumnIndex);            	
		if(columnName != null){	//skip columns with empty title
			ColumnAttr columnAttr = attrMap.get(columnName);
	    	if(columnAttr.isColNumCalculatable()){
	    		Double count = funcNumber(mySheet, myWorkBook,curColumnIndex, rowResult, "—◊®“");
	    		if(count > MIN_MAX_LOWER){
		    		Double min = funcNumber(mySheet, myWorkBook,curColumnIndex, rowResult, "Ã»Õ");
		    		Double max = funcNumber(mySheet, myWorkBook,curColumnIndex, rowResult, "Ã¿ —");
		    		setRemoveMinMax.add(e)
	    		}
	    	}	
	    }
	}
	NavigableSet<Integer> rmMinMax = ((TreeSet)setRemoveMinMax).descendingSet();
	for (Integer rowInd : rmMinMax) {
		System.out.println("DEBUG: 101 : will be removed rowInd =" + rowInd);
		removeRow(mySheet, rowInd);
	}
}
*/


/*        
SortedSet<Integer> setRemoveMinMax = new TreeSet<Integer>();
if(REMOVE_MIN_MAX){
	for (Cell cell : titleRow){
		int curColumnIndex = cell.getColumnIndex();
		String columnName = indexMap.get(curColumnIndex);            	
		if(columnName != null){	//skip columns with empty title
			ColumnAttr columnAttr = attrMap.get(columnName);
	    	if(columnAttr.isColNumCalculatable()){
	    		System.out.println("DEBUG: 100 :  columnName =" + columnName);
	    		setRemoveMinMax.add(columnAttr.getMinValue().getRowNum());
	    		setRemoveMinMax.add(columnAttr.getMaxValue().getRowNum());
	    	}
	    }
	}
	NavigableSet<Integer> rmMinMax = ((TreeSet)setRemoveMinMax).descendingSet();
	for (Integer rowInd : rmMinMax) {
		System.out.println("DEBUG: 101 : will be removed rowInd =" + rowInd);
		removeRow(mySheet, rowInd);
	}
}
*/        



/*
private static void procRow(Cell cell) {
	if(cell.getCellType() == HSSFCell.CELL_TYPE_STRING){
	    String cellValue = cell.getStringCellValue();
	    //System.out.print("cellValue : " + cellValue);
	    System.out.print(cellValue + " ");
	}
	
//	if(cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC){
//		Date birthdate = cell.getDateCellValue();
//		System.out.print("birthdate :" + birthdate);
//	}
}
*/	



/*	
private static double funcNumber(XSSFSheet mySheet, XSSFWorkbook myWorkBook, int curColumnIndex, XSSFRow rowResult, String formula){
	//XSSFRow rowResult = mySheet.createRow(lastRowNum + SHIFT_FORMULA_ROW);
	//Cell cell = rowResult.createCell(1);
	Cell cell = rowResult.createCell(curColumnIndex);
	//System.out.println("getCellName(beginRange)=" + getCellName(beginRange));
	//System.out.println("getCellName(endRange)=" + getCellName(endRange));
	String range = getCellName(curColumnIndex,FIRST_ROW) + ":" + getCellName(curColumnIndex,lastRowNum);	//C2:C10
	System.out.println("Range=" + range);
	
	cell.setCellType(XSSFCell.CELL_TYPE_FORMULA);
	String form = formula + "(" + range + ")";
	cell.setCellFormula(form);	    
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
*/	





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