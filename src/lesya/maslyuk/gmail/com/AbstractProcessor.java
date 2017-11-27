package lesya.maslyuk.gmail.com;

import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import lesya.maslyuk.gmail.com.model.ColumnAttr;
import lesya.maslyuk.gmail.com.model.ColumnType;
import lesya.maslyuk.gmail.com.utils.MyUtils;

public abstract class AbstractProcessor {
	
	protected void process(XSSFSheet mySheet, Map<Integer,String> indexMap, Map<String,ColumnAttr> attrMap){
        for (Row row : mySheet) {
        	int zeroLine = row.getRowNum();
        	if (zeroLine == 0)
        		continue;
        	
        	for (Cell cell : row) {                	
    			String columnName = indexMap.get(cell.getColumnIndex());
    			if(columnName != null){	//skip columns with empty title
    		    	ColumnAttr columnAttr = attrMap.get(columnName);
		    		if(!MyUtils.isCellEmpty(cell)){
				    	//fillMinMax(row, cell, columnAttr);
		    			processConcret(row, cell, columnAttr);
		    		}
    			}
    		}            
        }
	}
	
	protected abstract void processConcret(Row row, Cell cell, ColumnAttr columnAttr);
	
}
