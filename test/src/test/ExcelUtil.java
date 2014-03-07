/*java-excel v1 branch test*/
package test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.regex.Pattern;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelUtil{
	public static void main(String[] argv) {
		try {
			ExcelUtil.exportExcelForStudent();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		System.out.println("哈哈成功");
	}

	public static void exportExcelForStudent() throws IOException {		//创建excel文件对象
		SimpleDateFormat sdf = new SimpleDateFormat("HH:mm:ss:SS");
		Long startTime = System.currentTimeMillis();
		Workbook[] wbs = new Workbook[] { new HSSFWorkbook(),new XSSFWorkbook() };
		for(int i=0; i<wbs.length; i++) {
			Workbook wb = wbs[i];
		//创建一个张表
		Sheet sheet = wb.createSheet();
		//创建第一行
		Row row = sheet.createRow(0);
	               //创建第二行
		Row row1 = sheet.createRow(1);
		// 文件头字体
		Font font0 = createFonts(wb, Font.BOLDWEIGHT_BOLD, "宋体", false,
				(short) 200);
		Font font1 = createFonts(wb, Font.BOLDWEIGHT_NORMAL, "宋体", false,
				(short) 200);
		CellStyle cellStyle = wb.createCellStyle();
		cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
		cellStyle.setVerticalAlignment(CellStyle.VERTICAL_BOTTOM);
		cellStyle.setFont(font0);
		// 合并第一行的单元格
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 1));
		//设置第一列的文字
		createCell(wb, row, 0, "总数", cellStyle);
		//合并第一行的2列以后到8列（不包含第二列）
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 2, 8));
		//设置第二列的文字
		createCell(wb, row, 2, "基本信息", cellStyle);
		//给第二行添加文本,重新设置字体
		cellStyle.setFont(font1);
		createCell(wb, row1, 0, "序号", cellStyle);
		createCell(wb, row1, 1, "版本", cellStyle);
		createCell(wb, row1, 2, "姓名", cellStyle);
		createCell(wb, row1, 3, "性别", cellStyle);
		createCell(wb, row1, 4, "年龄", cellStyle);
		createCell(wb, row1, 5, "年级", cellStyle);
		createCell(wb, row1, 6, "学校", cellStyle);
		createCell(wb, row1, 7, "父母名称", cellStyle);
		createCell(wb, row1, 8, "籍贯", cellStyle);
		createCell(wb, row1, 9, "联系方式", cellStyle);
		//第三行表示
		int l = 2;
		//这里将学员的信心存入到表格中		
		for (int j = 0; j < 10000; j++) {
			//创建一行
			Row rowData = sheet.createRow(l++);
			//Student stu = studentList.get(i);
			createCell(wb, rowData, 0, String.valueOf(j + 1), cellStyle);
			createCell(wb, rowData, 1, "3.0", cellStyle);
			createCell(wb, rowData, 2, "陈明龙", cellStyle);
			createCell(wb, rowData, 3, "男", cellStyle);
			 createCell(wb, rowData, 4, "23", cellStyle);
			createCell(wb, rowData, 5, "一年级", cellStyle);
			createCell(wb, rowData, 6, "浙江大学", cellStyle);
			createCell(wb, rowData, 7, "陈太平", cellStyle); 
			createCell(wb, rowData, 8, "安徽省全椒县大墅镇", cellStyle);
			createCell(wb, rowData, 9, "15271815754", cellStyle);
	
		}
		// Save
		   String filename = "D:\\workbook.xls";
		   if(wb instanceof XSSFWorkbook) {
		     filename = filename + "x";
		   }
		 
		   FileOutputStream out = new FileOutputStream(filename);
		   wb.write(out);
		   out.close();
		   Long endTime = System.currentTimeMillis();
	        System.out.println(filename+"用时：" + sdf.format(new Date(endTime - startTime)));
		}
	}	

/**
	 * 创建单元格并设置样式,值
	 * 
	 * @param wb
	 * @param row
	 * @param column
	 * @param
	 * @param
	 * @param value
	 */
	public static void createCell(Workbook wb, Row row, int column,
			String value, CellStyle cellStyle) {
		Cell cell = row.createCell(column);
		cell.setCellValue(value);
		cell.setCellStyle(cellStyle);
	}

	/**
	 * 设置字体
	 * 
	 * @param wb
	 * @return
	 */
	public static Font createFonts(Workbook wb, short bold, String fontName,
			boolean isItalic, short hight) {
		Font font = wb.createFont();
		font.setFontName(fontName);
		font.setBoldweight(bold);
		font.setItalic(isItalic);
		font.setFontHeight(hight);
		return font;
	}

	/**
	 * 判断是否为数字
	 * 
	 * @param str
	 * @return
	 */
	public static boolean isNumeric(String str) {
		if (str == null || "".equals(str.trim()) || str.length() > 10)
			return false;
		Pattern pattern = Pattern.compile("^0|[1-9]\\d*(\\.\\d+)?$");
		return pattern.matcher(str).matches();
	}

}
