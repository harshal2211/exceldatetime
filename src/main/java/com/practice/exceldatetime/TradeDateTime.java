package com.practice.exceldatetime;

import java.io.File;
import java.io.IOException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;
import java.util.TimeZone;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class TradeDateTime {
	
	public enum DateTimeType{
		DATE,
		TIME,
		DATETIME
	}

	public static final String XLSX_PATH="C:\\Users\\dell\\Desktop\\dateformats.xlsx";
	public static int[] styles = {DateFormat.FULL, DateFormat.LONG, DateFormat.MEDIUM, DateFormat.SHORT};
		
	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		Workbook workbook = WorkbookFactory.create(new File(XLSX_PATH));
		DataFormatter dataFormatter = new DataFormatter();
        Sheet sheet = workbook.getSheetAt(2);
        for (Row row: sheet) {
        	if(row.getRowNum()==0)
        		continue;
        	//for(Cell cell: row) {
        		String tz = dataFormatter.formatCellValue(row.getCell(0));
        		String lcle = dataFormatter.formatCellValue(row.getCell(1));
        		String trdDateString = dataFormatter.formatCellValue(row.getCell(2));       		
        		String trdTimeString = dataFormatter.formatCellValue(row.getCell(3));
        		String tdrContactTime = dataFormatter.formatCellValue(row.getCell(4));
        		String clntOrderTime = dataFormatter.formatCellValue(row.getCell(5));
        		TimeZone timezone = TimeZone.getTimeZone(tz);
        		Locale locale = new Locale(lcle);
        		System.out.println("**********"+row.getRowNum()+"************");
        		System.out.println("Timezone "+ timezone);
        		System.out.println("Locale "+ locale);
        		//System.out.println("Trade Date in String "+ trdDateString + " \t cell format "+ getCellFormat(row.getCell(2)) + " \t Date Object "+ 
        		//		getDateCellValue(row.getCell(2), DateTimeType.DATE, locale));
        		//System.out.println("Trade Time in String "+ trdTimeString + " \t cell format "+ getCellFormat(row.getCell(3)) + " \t Date Object "+ 
                //		getDateCellValue(row.getCell(3), DateTimeType.TIME, locale));
        		System.out.println("Trader Contact Time in String "+ tdrContactTime + " \t cell format "+ getCellFormat(row.getCell(4)) + " \t Date Object "+ 
                		getDateCellValue(row.getCell(4), DateTimeType.DATETIME, locale));
        		System.out.println("Client Order Time in String "+ clntOrderTime + " \t cell format "+ getCellFormat(row.getCell(5)) + " \t Date Object "+ 
                		getDateCellValue(row.getCell(5), DateTimeType.DATETIME, locale));
        	//}        	
        }
	}
	
	public static String getCellFormat(Cell cell) {
		String cellFormat = null;
		if(cell!= null) {
			CellStyle trdDateStyle = cell.getCellStyle();
			short trdDateFormatIndex = trdDateStyle.getDataFormat();
			cellFormat = BuiltinFormats.getBuiltinFormat(trdDateFormatIndex);
		}
		return cellFormat;
	}
	
	public static Date getDateCellValue(Cell cell, DateTimeType dateTimeType, Locale userLocale) {
		Date date = null;
		if(cell != null) {
			if(cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
				return cell.getDateCellValue();
			} else {
				DataFormatter dataFormatter = new DataFormatter();
				String dateCellValue = dataFormatter.formatCellValue(cell);
				dateCellValue = dateCellValue.replaceAll("-", "/");
				dateCellValue = dateCellValue.trim();
				date = parseStringAsDate(dateCellValue, dateTimeType, userLocale);
				if(date == null) {
					date = parseStringAsDateUsingCellFormat(dateCellValue, getCellFormat(cell));
				}
			}
		}
		return null;
	}
	
	private static Date parseStringAsDateUsingCellFormat(String dateCellValue, String cellFormat) {
		SimpleDateFormat sdf = new SimpleDateFormat(cellFormat);
		Date date = null;
		try {
			date = sdf.parse(dateCellValue);
		} catch (ParseException e) {
			System.out.println("parse error with format "+ cellFormat);
		}
		return date;
	
	}

	public static Date parseStringAsDate(String value, DateTimeType dateTimeType, Locale userLocale) {
		Date date = null;
		switch(dateTimeType) {
		case DATE:date = formatDate(value, userLocale);
			break;
		case DATETIME:date = formatDateTime(value, userLocale);
			break;
		case TIME:date = formatTime(value, userLocale);
			break;
		default:
			break;
		}
		return date;
		
	}
	
	private static Date formatDateTime(String field, Locale userLocale) {
		for(int k : styles) {
			DateFormat format = DateFormat.getDateTimeInstance(styles[k], styles[k], Locale.getDefault());
			format.setTimeZone(TimeZone.getDefault());
			try {
				//System.out.println("style "+k+ " pattern"+ ((SimpleDateFormat)format).toPattern());
				Date date = format.parse(field);
				System.out.println("style "+ k +"date "+date+ " pattern "+ ((SimpleDateFormat)format).toPattern());
				return date;
			} catch (ParseException e) {
				//System.out.println("parsing failed with styles" + styles[k]);
			}
		}
		return null;
	}
	
	private static Date formatDate(String field, Locale userLocale) {
		for(int k : styles) {
			DateFormat format = DateFormat.getDateInstance(styles[k], userLocale);
			format.setTimeZone(TimeZone.getDefault());
			try {
				Date date = format.parse(field);
				System.out.println("style "+ k +"date "+date+ " pattern "+ ((SimpleDateFormat)format).toPattern());
				return date;
			} catch (ParseException e) {
				//System.out.println("parsing failed with styles" + styles[k]);
			}
		}
		return null;
	}
	
	private static Date formatTime(String field, Locale userLocale) {
		for(int k : styles) {
			DateFormat format = DateFormat.getTimeInstance(styles[k], userLocale);
			format.setTimeZone(TimeZone.getDefault());
			//System.out.println(" pattern "+ ((SimpleDateFormat)format).toPattern());
			try {
				Date date = format.parse(field);
				System.out.println("style "+ k +"date "+date+ " pattern "+ ((SimpleDateFormat)format).toPattern());
				return date;
			} catch (ParseException e) {
				//System.out.println("parsing failed with styles" + styles[k]);
			}
		}
		return null;
	}
}
