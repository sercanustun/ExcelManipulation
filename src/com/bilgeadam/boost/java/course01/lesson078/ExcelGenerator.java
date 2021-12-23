package com.bilgeadam.boost.java.course01.lesson078;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.Date;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

public class ExcelGenerator {
	private static final String EXCEL_FILE = "C:\\Users\\sercu\\SERCAN\\DERSICI\\My Excel.xlsx";
	
	public static void main(String[] args) throws IOException {
		
		generate();
	}
	
	private static void generate() throws IOException {
		
		Map<String, ExcelRow> data = new TreeMap<>();
		data.put("1", new ExcelRow(1, "Babur", "2021-12-19", 12.3f));
		data.put("2", new ExcelRow(2, "Cafer", "2020-12-19", 178.8f));
		data.put("3", new ExcelRow(3, "Narin", "2019-12-19", 23f));
		data.put("4", new ExcelRow(4, "Tezer", "2018-12-19", 48.33465f));
		data.put("5", new ExcelRow(5, "Selim", "2017-12-19", 767f));
		data.put("6", new ExcelRow(6, "Alev", "2016-12-19", 34523.234567f));
		
		XSSFWorkbook wb = new XSSFWorkbook(); // �al���lacak workbokk'u yarat
		XSSFSheet sheet = wb.createSheet("Benim Sheetim");
		sheet.setTabColor(new XSSFColor(IndexedColors.CORNFLOWER_BLUE));
		sheet.setColumnWidth(3, 15 * 256);
		sheet.autoSizeColumn(2);
		
		XSSFRow row;
		int rowCnt = 0;
		String[] header = { "Id", "Isim", "Kayit Tarihi", "Boy" };
		CellStyle headerStyle = wb.createCellStyle();
		headerStyle.setBorderBottom(BorderStyle.THICK);
		headerStyle.setBorderTop(BorderStyle.THICK);
		headerStyle.setBorderRight(BorderStyle.THICK);
		headerStyle.setBorderLeft(BorderStyle.THICK);
		headerStyle.setFillForegroundColor((short) 200);
		headerStyle.setFillPattern(FillPatternType.LEAST_DOTS);
		headerStyle.setAlignment(HorizontalAlignment.CENTER);
		XSSFFont font = wb.createFont();
		font.setBold(true);
		font.setFontHeightInPoints((short) 15);
		headerStyle.setFont(font);
		
		sheet.createRow(rowCnt++);
		row = sheet.createRow(rowCnt++);
		int colCnt = 0;
		row.createCell(colCnt++);
		for (String string : header) {
			Cell cell = row.createCell(colCnt++);
			cell.setCellValue(string);
			cell.setCellStyle(headerStyle);
		}
		
		CellStyle cellStyle = wb.createCellStyle();
		CreationHelper createHelper = wb.getCreationHelper();
		cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("[$-tr-TR]d mmm yyyy;@"));
		
		Font fnt = wb.createFont();
		fnt.setColor(new XSSFColor(IndexedColors.LAVENDER).getIndex());
		fnt.setBold(true);
		
		Set<String> keys = data.keySet();
		for (String key : keys) {
			row = sheet.createRow(rowCnt++);
			ExcelRow rowData = data.get(key);
			short cellCnt = 0;
			row.createCell(cellCnt++);
			Cell cell1 = row.createCell(cellCnt++);
			cell1.setCellValue(rowData.getNumber());
			Cell cell2 = row.createCell(cellCnt++);
			cell2.setCellValue(rowData.getName());
			Cell cell3 = row.createCell(cellCnt++);
			Date dt = Date.from(rowData.getDate().atStartOfDay(ZoneId.systemDefault()).toInstant());
			cell3.setCellValue(dt);
			cellStyle.setFont(fnt);
			cell3.setCellStyle(cellStyle);
			
			Cell cell4 = row.createCell(cellCnt++);
			cell4.setCellValue(rowData.height);
		}
		
		FileOutputStream fos = new FileOutputStream(new File(ExcelGenerator.EXCEL_FILE));
		wb.write(fos);
		wb.close();
		fos.close();
	}
	
	private static class ExcelRow {
		private int number;
		private String name;
		private LocalDate date;
		private float height;
		
		public ExcelRow(int number, String name, String date, float height) {
			super();
			this.number = number;
			this.name = name;
			this.date = LocalDate.parse(date);
			this.height = height;
		}
		
		public int getNumber() {
			return number;
		}
		
		public String getName() {
			return name;
		}
		
		public LocalDate getDate() {
			return date;
		}
	}
}