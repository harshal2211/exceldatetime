package com.practice.exceldatetime;

import java.io.File;
import java.io.IOException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.ZoneOffset;
import java.time.ZonedDateTime;
import java.util.Calendar;
import java.util.Date;
import java.util.Locale;
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

public class ReadXls2 {
	
public static final String XLSX_PATH="C:\\Users\\dell\\Desktop\\dateformats.xlsx";
public static int[] styles = {DateFormat.FULL, DateFormat.LONG, DateFormat.MEDIUM, DateFormat.SHORT};
	
	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		Workbook workbook = WorkbookFactory.create(new File(XLSX_PATH));
        for(Sheet sheet: workbook) {
            System.out.println("=> " + sheet.getSheetName());
        }
        
        DataFormatter dataFormatter = new DataFormatter();
        Sheet sheet = workbook.getSheetAt(0);
        
        /*for (Row row: sheet) {
        	if(row.getRowNum() == 0) 
        		continue;
        	String dateTimeCase = null;
        	String field = null;
        	Date excelDate = null;
            for(Cell cell: row) {
            	if(cell.getColumnIndex() == 0) {
            		System.out.print("\n Format "+ dataFormatter.formatCellValue(cell));
            		continue;
            	}
            	if(cell.getColumnIndex() == 1) {
            		dateTimeCase = dataFormatter.formatCellValue(cell);
            		continue;
            	}
            	if(cell.getColumnIndex() == 2) {
            		field = dataFormatter.formatCellValue(cell);
            		if(cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
            			excelDate = cell.getDateCellValue();
            		}
            		continue;
            	}	
            }
            parseStringAsDate(excelDate, field, dateTimeCase);
        }*/
        
        sheet = workbook.getSheetAt(1);
        for (Row row: sheet) {
        	System.out.println(row.getRowNum());
        	Date excelDate = null;
        	Date excelTime = null;
        	 for(Cell cell: row) {
        		 if(cell.getColumnIndex() == 0) {
        			 String field = dataFormatter.formatCellValue(cell);
        			 if(cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
        				System.out.println(field);
             			excelDate = cell.getDateCellValue();
             		} else {
             			excelDate = parseStringAsDate(excelDate, field, "date");
             		}
        		 }
        		 if(cell.getColumnIndex() == 1) {
        			 String field = dataFormatter.formatCellValue(cell);
        			 if(cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
        				System.out.println(field);
             			excelTime = cell.getDateCellValue();
             		} else {
             			excelTime = parseStringAsDate(excelDate, field, "time");
             		}
        		 }
        	 }
        	 System.out.println(getCombinedDate(excelDate, excelTime));
        }
	}
	
	/*private static void parseStringAsLocalDateTime(String value) {
		FormatStyle[] styles = {FormatStyle.FULL, FormatStyle.LONG, FormatStyle.MEDIUM, FormatStyle.SHORT};
		DateTimeFormatter formatter;
		LocalDateTime result = null;
		for (int k = 0; k < styles.length; k++) {
	         formatter = DateTimeFormatter.ofLocalizedDateTime(styles[k]);
	         try {
	        	 String pattern = DateTimeFormatterBuilder.getLocalizedDateTimePattern(styles[k], styles[k], IsoChronology.INSTANCE, Locale.getDefault());
	        	 System.out.print("\n style "+ styles[k] + " pattern "+ pattern);
	        	 result = (LocalDateTime) formatter.parse(value);
	        	 System.out.print("\n"+result);
	        	 return;
	         }catch(DateTimeParseException ex) {
	        	 System.out.print("\n error using format style "+styles[k]);
	         }
	      }
		LocalDateTime.parse(text, formatter)
	}*/
	
	private static Date parseStringAsDate(Date excelDate, String field, String dateTimeCase) {
		field = field.replaceAll("-", "/");
		field = field.trim();
		System.out.println("Value "+ field);
		System.out.println("Excel Date "+ excelDate);
		System.out.println("Date Time case "+ dateTimeCase);
		Date date = null;
		switch(dateTimeCase) 
		{
		case "datetime": date = formatDateTime(field);
						LocalDateTime localDateTime = convertDateToLocalDateTime(date!=null? date: excelDate);
						System.out.println("Local Date time "+ localDateTime);
						break;
		case "date": date = formatDate(field);
						LocalDate localDate = convertDateToLocalDate(date!=null? date: excelDate);
						System.out.println("Local Date "+ localDate);
					break;
		case "time": date = formatTime(field);
					break;
		}
		return date;
	}
	
	private static LocalDate convertDateToLocalDate(Date date) {
		LocalDate ld = date.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
		return ld;
	}

	private static LocalDateTime convertDateToLocalDateTime(Date date) {
		LocalDateTime ldt = date.toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime();
		System.out.println(ldt);
		LocalDateTime eldt = date.toInstant().atZone(ZoneId.of("Europe/Rome")).toLocalDateTime();
		System.out.println(eldt);
		ZonedDateTime zdt = ZonedDateTime.of(ldt, ZoneId.of("Europe/Rome"));
		System.out.println(zdt);
		System.out.println(zdt.getOffset());
		//ZonedDateTime utcZoned = ZonedDateTime.of(LocalDate.now().atTime(11, 30), ZoneOffset.UTC);
		//ZoneId swissZone = ZoneId.of("Europe/Zurich");
		//ZonedDateTime swissZoned = utcZoned.withZoneSameInstant(swissZone);
		//LocalDateTime swissLocal = swissZoned.toLocalDateTime();
		return ldt;
	}

	private static Date formatDateTime(String field) {
		for(int k : styles) {
			DateFormat format = DateFormat.getDateTimeInstance(styles[k], styles[k], Locale.getDefault());
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
	
	private static Date formatDate(String field) {
		for(int k : styles) {
			DateFormat format = DateFormat.getDateInstance(styles[k], Locale.getDefault());
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
	
	private static Date formatTime(String field) {
		for(int k : styles) {
			DateFormat format = DateFormat.getTimeInstance(styles[k], Locale.getDefault());
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
	
	private static Date getCombinedDate(Date date, Date time) {
		System.out.println("date part "+ date);
		System.out.println("time part "+ time);
		Calendar dateCalendar = Calendar.getInstance();
		dateCalendar.setTime(date);
		Calendar timeCalendar = Calendar.getInstance();
		timeCalendar.setTime(time);
		dateCalendar.set(Calendar.HOUR_OF_DAY, timeCalendar.get(Calendar.HOUR_OF_DAY));
		dateCalendar.set(Calendar.MINUTE, timeCalendar.get(Calendar.MINUTE));
		dateCalendar.set(Calendar.SECOND, timeCalendar.get(Calendar.SECOND));
		
		Date combinedDate = dateCalendar.getTime();
		LocalDateTime ldt = combinedDate.toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime();
		System.out.println("local date time "+ ldt);
		ZonedDateTime zdt = ZonedDateTime.of(ldt, TimeZone.getTimeZone("Europe/Rome").toZoneId());
		System.out.println(zdt.getOffset());
		return combinedDate;
	}
	
	
}
