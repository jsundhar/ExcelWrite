package com.waitrose.data;

import java.io.FileOutputStream;
import java.time.LocalDateTime;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;


public class ExcelDataManipulator {

	public static void main(String[] args) {
		SXSSFWorkbook wb =  new SXSSFWorkbook();
	    String safeName = WorkbookUtil.createSafeSheetName("[Waitrose testcases]"); 
	    SXSSFSheet sheet = wb.createSheet(safeName);

	    FileOutputStream fileOut;
	    wb.setCompressTempFiles(true);
		try {
	        wb.setCompressTempFiles(true);
	        SXSSFSheet sh = (SXSSFSheet) wb.getSheetAt(0);
	        sh.setRandomAccessWindowSize(100);
	        
	        for(int rownum = 0; rownum < 100000; rownum++){
	        	Row row = sh.createRow(rownum);
	        	Cell cell1 = row.createCell(0);
	        	Cell cell2 = row.createCell(1);
	        	cell1.setCellValue(rownum+1);
	        	cell2.setCellValue("HELLO");

	        }    
	        String timestamp = LocalDateTime.now ( ).toString ().replace ( "T", " " );
	        fileOut = new FileOutputStream("WaitroseExcel"+timestamp+".xlsx");
			wb.write(fileOut);
		    fileOut.close();
		    wb.dispose();
		} catch (Exception e) {
			e.printStackTrace();
		}

	}
	
}
