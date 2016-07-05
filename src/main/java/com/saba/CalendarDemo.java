package com.saba;

/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */


import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * A  monthly calendar created using Apache POI. Each month is on a separate sheet.
 * <pre>
 * Usage:
 * CalendarDemo -xls|xlsx <year>
 * </pre>
 *
 * @author Yegor Kozlov
 */
public class CalendarDemo {

    private static final String[] awardHeaders = {
        "Contact Details", "Awarded Bid Details", "Product Details"};
    
    private static final String[] contactDetails = {
        "Health System", "Contact Name", "Title", "Email","Phone"};
    
    private static final String[] awardedBidDetails = {
        "Requesting Department", "Facilities", "Buyer Part Number", "Description", "Contract Effective Date", "Contract End Date",
        "Total Value"};
    
    private static final String[] productDetailsHeaders = {
        "Product/Service Name", "UNSPSC", "Unit of Measure", "Quantity of Each", "Distributor(s)", 
        "Unit Price", "Discounted Unit Price", "Quantity Needed (per Yr.)", "Total Price"};

    public static void prepareXLSDynamicValues(Map<String, Object> data) {
    	data.put(contactDetails[0], "Columbus Health care Hospital");
    	data.put(contactDetails[1], "Sabari nathan");
    	data.put(contactDetails[2], "Purchasing Unit Head");
    	data.put(contactDetails[3], "sabar@yopmail.com");
    	data.put(contactDetails[4], "+91-908-765-4321");
    	
    	data.put(awardedBidDetails[0], "Cardiology, Administration, Admissions, Behavioral Health, Bloodborne Pathogen, Safety Device.");
    	data.put(awardedBidDetails[1], "PeachCare Medical Center, Unicol Country Memorial Hospital, Brandon Ambulatory Surgery Center");
    	data.put(awardedBidDetails[2], new Date());
    	data.put(awardedBidDetails[3],  new Date());
    	data.put(awardedBidDetails[4], new Double("1234.00"));
    	
    	Map<String, Object[]> productDetailsMap = new HashMap<String, Object[]>();
    	productDetailsMap.put("0", new Object[] {"Cardiology, Administration, Safety Device ",
    			"91101501-Health or fitness clubs", "EA","100","flipfort logic tech","10","9","45","450"});
    	productDetailsMap.put("1", new Object[] {"Admissions, Behavioral Health, Bloodborne Pathogen, Safety Device ",
    			"91101501-fitness clubs", "EA","100","lupanisa","10","9","45","450"});
    	productDetailsMap.put("2", new Object[] {"Behavioral Health, Bloodborne Pathogen, Safety Device ",
    			"91101501-Health or fitness clubs", "EA","100","Tech mahe logistics","10","9","45","450"});
    	productDetailsMap.put("3", new Object[] {"Safety Device ",
    			"91101501-Health or fitness clubs", "EA","100"," first choice","10","9","45","450"});
    	productDetailsMap.put("4", new Object[] {"Bloodborne Pathogen, Safety Device ",
    			"91101598- lubs", "EA","100"," Dlhivery","10","9","45","450"});
    	data.put(awardHeaders[2], productDetailsMap);
    }    
    
    public static void main(String[] args) throws Exception {
    	
    	Map<String, Object> data = new HashMap<String, Object>();
    	prepareXLSDynamicValues(data);
    	
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Awarded Bid & Contact Details");

        Map<String, CellStyle> styles = createStyles(workbook);
        sheet.setPrintGridlines(false);
        sheet.setDisplayGridlines(false);
        PrintSetup printSetup = sheet.getPrintSetup();
        printSetup.setLandscape(true);
        sheet.setFitToPage(true);
        sheet.setHorizontallyCenter(true);
        
        setupColumnWidthForEachFields(sheet);
        
        //preparing the contact & details table along with data 
        prepareContactDetailsTableAndData(data, sheet, styles);
       
        int contactdetrow = contactDetails.length+2 ;
        //preparing the award bid & details table along with data 
        prepareAwardBidDetailsTableAndData(data, sheet, styles, contactdetrow);
        
        int awardDetailsRow = (contactDetails.length+awardedBidDetails.length+4);
        //preparing the product & details table 
        prepareProductDetailsTable(workbook, sheet, styles, awardDetailsRow);
        //preparing the product & details table data 
        prepareProductDetailsTableData(data, sheet, styles, awardDetailsRow);
        
        FileOutputStream out = new FileOutputStream("award_bid.xlsx");
        workbook.write(out);
        out.close();
        
    }


	private static void prepareContactDetailsTableAndData(
			Map<String, Object> data, XSSFSheet sheet,
			Map<String, CellStyle> styles) {
		XSSFRow titleRow = sheet.createRow(0);
        titleRow.setHeightInPoints(16);
        for (int i = 0; i <= 1; i++) {
            titleRow.createCell(i).setCellStyle(styles.get("title"));
        }
        XSSFCell titleCell = titleRow.getCell(0);
        titleCell.setCellValue(awardHeaders[0].toString());
        sheet.addMergedRegion(CellRangeAddress.valueOf("$A$1:$B$1"));
        
        for (int i = 0; i < contactDetails.length; i++) {
        	 XSSFRow row = sheet.createRow(i+1);
        	 XSSFCell cell = row.createCell(0);
             cell.setCellValue(contactDetails[i].toString());
             cell.setCellStyle(styles.get("item_left"));
             cell = row.createCell(1);
             populateDynamicObject(cell, data.get(contactDetails[i]));
             cell.setCellStyle(styles.get("item_right"));
        }
	}


	private static void prepareAwardBidDetailsTableAndData(
			Map<String, Object> data, XSSFSheet sheet,
			Map<String, CellStyle> styles, int contactdetrow) {
		XSSFRow titleRow;
		XSSFCell titleCell;
		titleRow = sheet.createRow(contactdetrow);
        titleRow.setHeightInPoints(16);
        for (int i = 0; i <= 1; i++) {
            titleRow.createCell(i).setCellStyle(styles.get("title"));
        }
        titleCell = titleRow.getCell(0);
        titleCell.setCellValue(awardHeaders[1].toString());
        sheet.addMergedRegion(CellRangeAddress.valueOf("$A$"+(contactdetrow+1)+":$B$"+(contactdetrow+1)));
        for (int i = 0; i < awardedBidDetails.length; i++) {
        	 XSSFRow row = sheet.createRow(contactdetrow+i+1);
        	 XSSFCell cell = row.createCell(0);
             cell.setCellValue(awardedBidDetails[i]);
             cell.setCellStyle(styles.get("item_left"));
             cell = row.createCell(1);
             populateDynamicObject(cell, data.get(awardedBidDetails[i]));
             cell.setCellStyle(styles.get("item_right"));
        }
	}


	private static void prepareProductDetailsTable(XSSFWorkbook workbook,
			XSSFSheet sheet, Map<String, CellStyle> styles, int awardDetailsRow) {
		XSSFRow titleRow;
		XSSFCell titleCell;
		titleRow = sheet.createRow(awardDetailsRow);
        titleRow.setHeightInPoints(16);
        for (int i = 0; i < productDetailsHeaders.length; i++) {
            titleRow.createCell(i).setCellStyle(styles.get("title"));
        }
        titleCell = titleRow.getCell(0);
        titleCell.setCellValue(awardHeaders[2].toString());
		String columId = productDetailsHeaders.length > 0
				&& productDetailsHeaders.length < 27 ? String
				.valueOf((char) (productDetailsHeaders.length + 'A' - 1)): null;
        String cellMergeRange = "$A$"+(awardDetailsRow+1)+":$"+columId+"$"+(awardDetailsRow+1);
        sheet.addMergedRegion(CellRangeAddress.valueOf(cellMergeRange));
        XSSFRow row = sheet.createRow(awardDetailsRow+1);
        for (int i = 0; i < productDetailsHeaders.length; i++) {
        	XSSFCell cell = row.createCell(i);
             cell.setCellValue(productDetailsHeaders[i]);
             //create header style for product Details table
     	     CellStyle headerStyle = createHeaderStyleForAward(workbook);
             cell.setCellStyle(headerStyle);
        }
	}


	private static void prepareProductDetailsTableData(
			Map<String, Object> data, XSSFSheet sheet,
			Map<String, CellStyle> styles, int awardDetailsRow) {
		if(data.containsKey(awardHeaders[2]) && null != data.get(awardHeaders[2])){
        	@SuppressWarnings("unchecked")
			Map<String, Object[]> productDetailsMap = (Map<String, Object[]>)data.get(awardHeaders[2]);
            Set<String> keyset = productDetailsMap.keySet();
    		int rownum = awardDetailsRow + 2;
            for (String key : keyset) {
    			try {
    			XSSFRow pDetailsRow = sheet.createRow(rownum++);
    			pDetailsRow.setHeightInPoints(12.75f);
    				Object[] objArr = productDetailsMap.get(key);
    				int cellnum = 0;
    				for (Object obj : objArr) {
    					XSSFCell cell = pDetailsRow.createCell(cellnum);
    					cell.setCellStyle(styles.get("item_right"));
    					//find and populate dynamic variable from object 
    					populateDynamicObject(cell, obj);
    					//increment the cell size
    					cellnum++;
    				}
    			} catch (Exception e) {
    				//logger.error("Error while preparing the product Details table in xls :" + e);
    				continue;
    			}
    		}
        }
	}

	private static void setupColumnWidthForEachFields(XSSFSheet sheet) {
		//Product/Service Name & Details Key(s)
        sheet.setColumnWidth(0, 180*33);
        //UNSPSC & Details value(s)
        sheet.setColumnWidth(1, 180*33);
        //Unit of Measure 
        sheet.setColumnWidth(2, 130*33);
        //Quantity of Each
        sheet.setColumnWidth(3, 130*33);
        //Distributor(s)
        sheet.setColumnWidth(4, 150*33);
        //Unit Price
        sheet.setColumnWidth(5, 100*33);
        //Discounted Unit Price
        sheet.setColumnWidth(6, 200*33);
        //Quantity Needed (per Yr.)
        sheet.setColumnWidth(7, 200*33);
        //Total Price
        sheet.setColumnWidth(8, 150*33);
	}

	private static void populateDynamicObject(Cell cell, Object obj) {
		if (obj instanceof Date) {
				cell.setCellValue((Date) obj);
			} else if (obj instanceof Boolean) {
				cell.setCellValue((Boolean) obj);
			} else if (obj instanceof String) {
				cell.setCellValue((String) obj);
			} else if (obj instanceof Double) {
				cell.setCellValue((Double) obj);
			}else if (obj instanceof Integer) {
				cell.setCellValue((Integer) obj);
			}
	}
    
    private static Map<String, CellStyle> createStyles(Workbook wb){
        Map<String, CellStyle> styles = new HashMap<String, CellStyle>();

        CellStyle style;
        Font titleFont = wb.createFont();
        titleFont.setFontHeightInPoints((short)14);
        titleFont.setFontName("Trebuchet MS");
        titleFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
        style = createBorderedStyle(wb);
        style.setFont(titleFont);
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        styles.put("title", style);
        

        Font itemFontLeft = wb.createFont();
        itemFontLeft.setFontHeightInPoints((short)11);
        itemFontLeft.setFontName("Trebuchet MS");
        itemFontLeft.setBoldweight(Font.BOLDWEIGHT_BOLD);
        style = createBorderedStyle(wb);
        style.setAlignment(CellStyle.ALIGN_LEFT);
        style.setFont(itemFontLeft);
        styles.put("item_left", style);

        Font itemFontRight = wb.createFont();
        itemFontRight.setFontHeightInPoints((short)10);
        itemFontRight.setFontName("Trebuchet MS");
        style = createBorderedStyle(wb);
        style.setAlignment(CellStyle.ALIGN_LEFT);
        style.setFont(itemFontRight);
        styles.put("item_right", style);
        
        return styles;
    }
    
 	/**
	 * createHeaderStyle : 
	 * Header row setting for sheet
	 * @param wb
	 * @return
	 */
	private static CellStyle createHeaderStyleForAward(Workbook wb) {
		CellStyle headerStyle;
		Font headerFont = wb.createFont();
		headerFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
		headerStyle = createBorderedStyle(wb);
		headerStyle.setAlignment(CellStyle.ALIGN_LEFT);
		headerStyle
				.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		headerStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		headerStyle.setFont(headerFont);
		return headerStyle;
	}

	/**
	 * CreateBorderedStyle : 
	 * left , right , bottom & top borderstyle setting for sheet
	 * @param wb
	 * @return
	 */
	private static CellStyle createBorderedStyle(Workbook wb) {
		CellStyle style = wb.createCellStyle();
		style.setBorderRight(CellStyle.BORDER_THIN);
		style.setRightBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderLeft(CellStyle.BORDER_THIN);
		style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderTop(CellStyle.BORDER_THIN);
		style.setTopBorderColor(IndexedColors.BLACK.getIndex());
		return style;
	}
	
}