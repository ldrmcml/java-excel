package test;

import org.apache.poi.ss.usermodel.*;

Workbook[] wbs = new Workbook[] { new HSSFWorkbook(), new XSSFWorkbook() };
for(int i=0; i<wbs.length; i++) {
   Workbook wb = wbs[i];
   CreationHelper createHelper = wb.getCreationHelper();

   // create a new sheet
   Sheet s = wb.createSheet();
   // declare a row object reference
   Row r = null;
   // declare a cell object reference
   Cell c = null;
   // create 2 cell styles
   CellStyle cs = wb.createCellStyle();
   CellStyle cs2 = wb.createCellStyle();
   DataFormat df = wb.createDataFormat();

   // create 2 fonts objects
   Font f = wb.createFont();
   Font f2 = wb.createFont();

   // Set font 1 to 12 point type, blue and bold
   f.setFontHeightInPoints((short) 12);
   f.setColor( IndexedColors.RED.getIndex() );
   f.setBoldweight(Font.BOLDWEIGHT_BOLD);

   // Set font 2 to 10 point type, red and bold
   f2.setFontHeightInPoints((short) 10);
   f2.setColor( IndexedColors.RED.getIndex() );
   f2.setBoldweight(Font.BOLDWEIGHT_BOLD);

   // Set cell style and formatting
   cs.setFont(f);
   cs.setDataFormat(df.getFormat("#,##0.0"));

   // Set the other cell style and formatting
   cs2.setBorderBottom(cs2.BORDER_THIN);
   cs2.setDataFormat(df.getFormat("text"));
   cs2.setFont(f2);


   // Define a few rows
   for(int rownum = 0; rownum < 30; rownum++) {
	   Row r = s.createRow(rownum);
	   for(int cellnum = 0; cellnum < 10; cellnum += 2) {
		   Cell c = r.createCell(cellnum);
		   Cell c2 = r.createCell(cellnum+1);
   
		   c.setCellValue((double)rownum + (cellnum/10));
		   c2.setCellValue(
		         createHelper.createRichTextString("Hello! " + cellnum)
		   );
	   }
   }
   
   // Save
   String filename = "workbook.xls";
   if(wb instanceof XSSFWorkbook) {
     filename = filename + "x";
   }
 
   FileOutputStream out = new FileOutputStream(filename);
   wb.write(out);
   out.close();
}