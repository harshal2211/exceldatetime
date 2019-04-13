package com.practice.exceldatetime;

import java.io.File;
import java.io.IOException;
import java.time.ZoneId;
import java.util.TimeZone;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ReadXls {
	
	public static final String XLSX_PATH="C:\\Users\\dell\\Documents\\test.xlsx";
	
	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		Workbook workbook = WorkbookFactory.create(new File(XLSX_PATH));
		
		System.out.println("Retrieving Sheets using for-each loop");
        for(Sheet sheet: workbook) {
            System.out.println("=> " + sheet.getSheetName());
        }
        
        DataFormatter dataFormatter = new DataFormatter();
        Sheet sheet = workbook.getSheetAt(0);
        
        System.out.println("\n\nIterating over Rows and Columns using for-each loop\n");
        for (Row row: sheet) {
            for(Cell cell: row) {
            	CellType cellType = cell.getCellType();
            	
            	switch(cellType) {
				case BLANK:
					break;
				case BOOLEAN:
					break;
				case ERROR:
					break;
				case FORMULA:
					break;
				case NUMERIC:
					if (DateUtil.isCellDateFormatted(cell)) {
						System.out.println("Date Cell Value :: "+cell.getDateCellValue()+"\t");
						System.out.println("Date Cell Value @ zone id of System ::"+cell.getDateCellValue()
						.toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime());
						System.out.println("Date Cell Java Date @ default timezone:: "+DateUtil.getJavaDate(cell.getNumericCellValue()));
						System.out.println("Date Cell Java Date with IST TZ :: "+
						DateUtil.getJavaDate(cell.getNumericCellValue(), TimeZone.getTimeZone("IST")));
						System.out.println("Date Cell Java Date @ zone id of System ::"+DateUtil.getJavaDate(cell.getNumericCellValue(), TimeZone.getTimeZone("IST"))
						.toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime());
						
					} else {
						//DataFormatter dataFormatter = new DataFormatter();
						/*System.out.println("YYYYYYYYY"+dataFormatter.formatCellValue(cell));
						System.out.println("cell value "+Double.valueOf(cell.getNumericCellValue()));
						BigDecimal bd = new BigDecimal(cell.getNumericCellValue());
						System.out.println("Big Decimal Value: "+bd);
						BigDecimal scaledBd = bd.setScale(10, BigDecimal.ROUND_HALF_UP);
						System.out.println("Big Decimal after scale 6 and round half up mode "+ scaledBd +"\t");
						System.out.println("Big Decimal to double "+ Double.valueOf(bd.toString()));
						System.out.println("Scaled Big Decimal to double "+ scaledBd.doubleValue());
						System.out.println("Big Decimal String to double "+ Double.parseDouble(bd.toString()));
						System.out.println("xxx" +new Double(cell.getNumericCellValue()));*/
						
						
						
					}
					break;
				case STRING:
					//System.out.print(cell.getStringCellValue()+"\t");
					break;
				case _NONE:
					break;
				default:
					break;
            	}
               /* String cellValue = dataFormatter.formatCellValue(cell);
                System.out.print(cellValue + "\t");*/
            }
            System.out.println();
        }
        
        workbook.close();
	}

}
