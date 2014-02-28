package test;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.TimeZone;
import java.util.logging.Level;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PoiReadExcel {
	/**
	 *基本的读取excel获取每行每列数据
	 * @throws InvalidFormatException 
	 */
    public static void main(String[] args) {
        SimpleDateFormat sdf = new SimpleDateFormat("HH:mm:ss:SS");
        TimeZone t = sdf.getTimeZone();
        t.setRawOffset(0);
        sdf.setTimeZone(t);
        Long startTime = System.currentTimeMillis();
        String fileName = "D:\\test2.xlsx";
        // 检测代码
        try {
            PoiReadExcel er = new PoiReadExcel();
            // 读取excel2007
            er.testPoiExcel2007(fileName);
        } catch (Exception e) {
            //Logger.getLogger(FastexcelReadExcel.class.getName()).log(Level.SEVERE, null, ex);
        	e.printStackTrace();
        }
        Long endTime = System.currentTimeMillis();
        System.out.println("用时：" + sdf.format(new Date(endTime - startTime)));
    }
	public void testPoiExcel2007(String strPath) throws IOException, InvalidFormatException {
        // 构造 XSSFWorkbook 对象，strPath 传入文件路径
		OPCPackage pkg = OPCPackage.open(strPath);
	    XSSFWorkbook xwb = new XSSFWorkbook(pkg);
        //XSSFWorkbook xwb = new XSSFWorkbook(strPath);
        // 读取第一章表格内容
        XSSFSheet sheet = xwb.getSheetAt(0);
        // 定义 row、cell
        XSSFRow row;
        String cell;
        // 循环输出表格中的内容
        for (int i = sheet.getFirstRowNum(); i < sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            /*System.out.println(row.getPhysicalNumberOfCells());
            for (int j = row.getFirstCellNum(); j < row.getPhysicalNumberOfCells(); j++) {
                // 通过 row.getCell(j).toString() 获取单元格内容，
                cell = row.getCell(j).toString();
                System.out.print(cell + "\t");
            }
            System.out.println("");System.out.println("abe:"+i);*/
            cell = row.getCell(0).toString();
            System.out.print(cell + "\t");
            cell = row.getCell(1).toString();
            System.out.print(cell + "\t");
            System.out.println();
        }      
	}
/*
依赖jar包：
dom4j-1.6.1.jar 
poi-ooxml-3.8-20120326.jar 
poi-3.8-20120326.jar 
xbean.jar 
poi-ooxml-schemas-3.8-20120326.jar */


}
