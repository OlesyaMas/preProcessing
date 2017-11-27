package lesya.maslyuk.gmail.com.model;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.CellStyle;
import lesya.maslyuk.gmail.com.utils.MyUtils;

public class ColumnAttr {
	public static final int FACTOR_CODING = 12;
	public static final int MINMAX_REMOVING_ROWS = 5;
	private static final double ORDINAL_FACTOR = 0.99;
	private int index;
	private ColumnType columnType = null;
	private Map<ColumnType,Integer> typeMap = new HashMap<ColumnType,Integer>();
	private double average;
	private CellStyle cellStyle;
	private MaxValue maxValue;
	private MinValue minValue;
	//int rowNum;
	//private Map<String, Integer> codingMap = new TreeMap<String,Integer>();
	
	private CodingMap codingMap = new CodingMap();
	
	public ColumnAttr(int index) {
		super();
		//this.index = index;
	}

	
	public CellStyle getCellStyle() {
		return cellStyle;
	}

	public void setCellStyle(CellStyle cellStyle) {
		this.cellStyle = cellStyle;
	}


//	public int getIndex() {
//		return index;
//	}
//	public void setIndex(int index) {
//		this.index = index;
//	}

	
	public double getAverage() {
		return average;
	}

	public void setAverage(double average) {
		this.average = average;
	}


	public MaxValue getMaxValue() {
		return maxValue;
	}

	public void setMaxValue(MaxValue maxValue) {
		this.maxValue = maxValue;
	}


	public MinValue getMinValue() {
		return minValue;
	}

	public void setMinValue(MinValue minValue) {
		this.minValue = minValue;
	}


	public ColumnType getColumnType() {
		//ColumnType result;
		if(columnType == null){
			columnType = ColumnType.BLANK;
			Map<ColumnType, Integer> sortedTypeMap = MyUtils.sortByValues((HashMap<ColumnType, Integer>)typeMap);
//			Set set2 = sortedTypeMap.entrySet();
//			Iterator iterator1 = set2.iterator();
//			while(iterator1.hasNext()){
//		    	Map.Entry mTypes = (Map.Entry)iterator1.next();
//		    	System.out.println("@@@@@@@@@@@" + mTypes.getKey() + "   " + mTypes.getValue());
//			}		
			
			Set<ColumnType> columnTypeSet = sortedTypeMap.keySet();
			Iterator iterator = columnTypeSet.iterator();
			while(iterator.hasNext()) {
		    	//Map.Entry<ColumnType, Integer> me = (Map.Entry)iterator.next();//java.lang.ClassCastException: ColumnType cannot be cast to java.util.Map$Entry
		    	ColumnType next = (ColumnType)iterator.next();
		    	//result = next.getKey();
		    	//System.out.print(me.getKey() + ": ");
		    	//System.out.println(me.getValue());
		    	if(particularlyFilled(columnTypeSet, next))
		    		next = (ColumnType)iterator.next();
		    	
		    	columnType = next;
		    	break;
		    }
		}
		return columnType; 
	}

	//TODO fix use MyUtils.isCellEmpty(cell)
	private boolean particularlyFilled(Set<ColumnType> columnTypeSet, ColumnType next) {
		return next == ColumnType.BLANK && columnTypeSet.size()>1;
	}

	public boolean isMinMaxRemovingApplicable() {
		return isColNumCalculatable() && typeMap.get(getColumnType()) >= MINMAX_REMOVING_ROWS;
	}
	
	//TODO use counts
	public boolean isColumnOrdinal(){
		boolean isOrdinal = false;
		int rowProceded = 0;
		if(getColumnType() == ColumnType.INTEGER){
			Set<Map.Entry<ColumnType,Integer>> typeSet = typeMap.entrySet();
			for (Iterator iterator = typeSet.iterator(); iterator.hasNext();) {
				Entry<ColumnType, Integer> entry = (Entry<ColumnType, Integer>) iterator.next();
				rowProceded = rowProceded + entry.getValue();
			}
		}
		isOrdinal = (rowProceded != 0 && typeMap.get(ColumnType.INTEGER) >= rowProceded * ORDINAL_FACTOR) ? true : false;
		return isOrdinal;
	} 
	
	//TODO add FACTOR
	public boolean isColumnCodeable(){
		boolean isCodeable = false;
		if(getColumnType() == ColumnType.TEXT){
			if(codingMap.size() > 1 && codingMap.size() <= FACTOR_CODING){
				//this.typeMap.get(ColumnType.TEXT) 0,3% 
				isCodeable = true;
			}
		}
		return isCodeable;
	}
	
	
	
	
	public Map<ColumnType, Integer> getTypeMap() {
		return typeMap;
	}
	public void setTypeMap(Map<ColumnType, Integer> typeMap) {
		this.typeMap = typeMap;
	}

	
	public CodingMap getCodingMap() {
		return codingMap;
	}


	public void setCodingMap(CodingMap codingMap) {
		this.codingMap = codingMap;
	}


	public boolean isColNumCalculatable() {
		//return columnAttr.getColumnType() == ColumnType.DECIMAL || (columnAttr.getColumnType() == ColumnType.INTEGER && !columnAttr.isColumnOrdinal());
		return getColumnType() == ColumnType.DECIMAL || (getColumnType() == ColumnType.INTEGER && !isColumnOrdinal());
	}
	
		
	@Override
	public String toString() {
		return "ColumnAttr [index=" + index + ", columnType=" + columnType + ", typeMap=" + typeMap + "]";
	}
}
