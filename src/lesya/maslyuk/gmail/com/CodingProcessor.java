package lesya.maslyuk.gmail.com;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import lesya.maslyuk.gmail.com.model.ColumnAttr;
import lesya.maslyuk.gmail.com.model.ColumnType;
import lesya.maslyuk.gmail.com.model.MaxValue;
import lesya.maslyuk.gmail.com.model.MinValue;

public class CodingProcessor extends AbstractProcessor {

	@Override
	protected void processConcret(Row row, Cell cell, ColumnAttr columnAttr) {
		System.out.println("DEBUG: CodingProcessor called");
		if(columnAttr.getColumnType() == ColumnType.TEXT){
			String textValue = cell.getStringCellValue() ;
			boolean loadingFactor = columnAttr.getCodingMap().size() <= ColumnAttr.FACTOR_CODING + 1;
			if(loadingFactor)
				columnAttr.getCodingMap().incValue(textValue);
		}
	}

}
