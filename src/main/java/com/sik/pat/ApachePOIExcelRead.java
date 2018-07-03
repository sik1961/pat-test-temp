package com.sik.pat;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.AbstractMap;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map.Entry;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ApachePOIExcelRead {

	private static final String FILE_NAME = "/home/sik/Documents/pattesttemp";
	private static final String FILE_EXTN = ".xlsx";
	private static final SimpleDateFormat SDF = new SimpleDateFormat("dd-MM-yyyy");
	private static final List<Integer> DATE_COLUMNS = new ArrayList<>();
	
	private static final int ROWS_TOP_MARGIN = 1;
	private static final int ROWS_PER_LABEL = 5;
	private static final Entry<Integer,Integer> LABEL_PAGE_LAYOUT = new AbstractMap.SimpleEntry<>(2,8);
	private static final Entry<Float,Float> PAGE_SIZE_POINTS = new AbstractMap.SimpleEntry<Float,Float>(595.0F,842.0F);
	private static final Float LABEL_CELL_WIDTH = (PAGE_SIZE_POINTS.getKey()/LABEL_PAGE_LAYOUT.getKey())/ROWS_PER_LABEL;
	private static final Float LABEL_CELL_HEIGHT = (PAGE_SIZE_POINTS.getValue()/LABEL_PAGE_LAYOUT.getValue())/ROWS_PER_LABEL;

	public static void main(String[] args) {

		try {

			FileInputStream excelFile = new FileInputStream(new File(FILE_NAME + FILE_EXTN));
			Workbook workbook = new XSSFWorkbook(excelFile);
			Sheet inputSheet = workbook.getSheetAt(0);
			Sheet labelSheet = workbook.createSheet("Labels");
			
			System.out.println("Cell: " + LABEL_CELL_WIDTH + "x" + LABEL_CELL_HEIGHT);
			
			//CellStyle labelCellStyle = workbook.createCellStyle();
			
			Iterator<Row> rowIterator = inputSheet.iterator();

			while (rowIterator.hasNext()) {

				//int 
				Row currentRow = rowIterator.next();
				
//				Row newLabelRow = getRow(labelSheet, getRowCell(currentRow.getRowNum()).getKey());
//				Cell newLabelCell = getCell(labelSheet, newLabelRow, getRowCell(currentRow.getRowNum()).getValue());
				
				Iterator<Cell> cellIterator = currentRow.iterator();

				while (cellIterator.hasNext()) {
					Cell currentCell = cellIterator.next();
					
					Row newLabelRow = getRow(labelSheet, getRowCell(currentRow.getRowNum()).getKey());
					Cell newLabelCell = getCell(labelSheet, newLabelRow, getRowCell(currentRow.getRowNum()).getValue());
					
					if (currentRow.getRowNum() == 0) {
						if (currentCell.getCellTypeEnum() == CellType.STRING
								&& currentCell.getStringCellValue().contains("Date")) {
							DATE_COLUMNS.add(currentCell.getColumnIndex());
						}
					}
					if (currentCell.getColumnIndex()==1) {
						newLabelCell.setCellValue(currentCell.getStringCellValue());
					}
					if (currentCell.getCellTypeEnum() == CellType.STRING) {
						System.out.print(currentCell.getStringCellValue() + ",");
					} else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
						if (isDate(currentCell)) {
							System.out.print(SDF.format(currentCell.getDateCellValue()) +",");
						} else {
							System.out.print(currentCell.getNumericCellValue() + ",");
						}
					}

				}
				System.out.println();

				FileOutputStream outputStream = new FileOutputStream(new File(FILE_NAME + "_Labels" + FILE_EXTN));
				workbook.write(outputStream);
				//workbook.close();
				
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	private static Cell getCell(Sheet sheet, Row row, Integer cellNbr) {
		Row newRow = getRow(sheet, row.getRowNum());
		Cell result = newRow.getCell(cellNbr);
		if (result == null) {
			result = newRow.createCell(cellNbr);
		}
		return result;
	}

	private static Row getRow(Sheet sheet, int row) {
		Row result = sheet.getRow(row);
		if (result == null) {
			result = sheet.createRow(row);
		} 
		return result;
	}

	private static Entry<Integer,Integer> getRowCell(int row) {
		int labelRow = ((row-1)/2)+1;
		int labelCell = (row%2);
		return new AbstractMap.SimpleEntry<Integer,Integer>(labelRow,labelCell);
	}

	private static boolean isDate(Cell currentCell) {
		return DATE_COLUMNS.contains(currentCell.getColumnIndex());
	}
}
