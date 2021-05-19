package com.bandar.excel.Entity;

import java.io.IOException;
import java.util.List;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CustomerExcelExporter {
	
	private XSSFWorkbook workbook;
	private XSSFSheet sheet;
	private List<Customer> listCustomer;
	
	public CustomerExcelExporter(List<Customer> listCustomer) {
		this.listCustomer = listCustomer;
		workbook = new XSSFWorkbook();
	}
	
	private void writeHeaderLine() {
        sheet = workbook.createSheet("Customer");
        // this to make the sheet from right to left
        sheet.getCTWorksheet().getSheetViews().getSheetViewArray(0).setRightToLeft(true);
        
        Row row = sheet.createRow(0);
         
        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setBold(true);
        font.setFontHeight(16);
        font.setColor(IndexedColors.RED.getIndex());
        
        style.setFont(font);
        style.getBorderBottom();
         
        createCell(row, 0, "الرقم", style);      
        createCell(row, 1, "الأسم", style);       
        createCell(row, 2, "الدولة", style);    
        createCell(row, 3, "العمر", style);
        createCell(row, 4, "الأيميل", style);   
    }
	
	private void createCell(Row row, int columnCount, Object value, CellStyle style) {
        sheet.autoSizeColumn(columnCount);
        Cell cell = row.createCell(columnCount);
        if (value instanceof Integer) {
            cell.setCellValue((Integer) value);
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        }else {
            cell.setCellValue((String) value);
        }
        cell.setCellStyle(style);
    }
	
	private void writeDataLines() {
        int rowCount = 1;
 
        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setFontHeight(14);
        style.setFont(font);
                 
        for (Customer customer : listCustomer) {
            Row row = sheet.createRow(rowCount++);
            int columnCount = 0;
             
            createCell(row, columnCount++, customer.getId(), style);
            createCell(row, columnCount++, customer.getName(), style);
            createCell(row, columnCount++, customer.getCountry(), style);
            createCell(row, columnCount++, customer.getAge(), style);
            createCell(row, columnCount++, customer.getEmail(), style);
             
        }
    }
	
	public void export(HttpServletResponse response) throws IOException {
        writeHeaderLine();
        writeDataLines();
         
        ServletOutputStream outputStream = response.getOutputStream();
        workbook.write(outputStream);
        workbook.close();
         
        outputStream.close();
         
    }
	
}
