package test;

import java.io.File;
import java.io.IOException;
import java.sql.Connection;
import java.sql.Date;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.TimeZone;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class ReadWriteExcelUtil_jxl {
	/**
	 * @param args
	 * @throws SQLException 
	 * @throws ClassNotFoundException 
	 */
	public static void main(String[] args) throws ClassNotFoundException, SQLException {
	    String fileName = "d:" + File.separator + "test.xls";//windows是\,unix是/
		ReadWriteExcelUtil_jxl util=new ReadWriteExcelUtil_jxl();
		String data=util.readExcel(fileName);			   
		System.out.println(data);
		String fileName1 = "d:" + File.separator + "abc.xls";
		//ReadWriteExcelUtil.writeExcel(fileName1,data);
	}

	/**
	 * 從excel文件中讀取所有的內容
	 * 
	 * @param file
	 *            excel文件
	 * @return excel文件的內容
	 * @throws ClassNotFoundException 
	 * @throws SQLException 
	 * @throws NumberFormatException 
	 */
	public String readExcel(String fileName) throws ClassNotFoundException, SQLException {
		Class.forName("com.mysql.jdbc.Driver");
    	Connection con = (Connection) DriverManager.getConnection("jdbc:mysql://" +
            "192.168.2.50:3306/excel2mysql", "root", "123");
    	// 关闭事务自动提交
    	con.setAutoCommit(false);
    	PreparedStatement pst = (PreparedStatement) con.prepareStatement("insert into test values (?,?)");
    	
    	String cellValue;
		StringBuffer sb = new StringBuffer();
		Workbook wb = null;
		try {
			// 构造Workbook（工作薄）对象
			wb = Workbook.getWorkbook(new File(fileName));
		} catch (BiffException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

		if (wb == null)
			return null;

		// 获得了Workbook对象之后，就可以通过它得到Sheet（工作表）对象了
		Sheet[] sheet = wb.getSheets();

		if (sheet != null && sheet.length > 0) {
			// 对每个工作表进行循环
			for (int i = 0; i < sheet.length; i++) {//工作表
				// 得到当前工作表的行数
				int rowNum = sheet[i].getRows();
				for (int j = 1; j < rowNum; j++) {//行
					// 得到当前行的所有单元格
					Cell[] cells = sheet[i].getRow(j);
					if (cells != null && cells.length > 0) {
						// 对每个单元格进行循环
/*						for (int k = 0; k < cells.length; k++) {//单元格cells.length的值有问题？？？
							// 读取当前单元格的值
							String cellValue = cells[k].getContents();
							
							try {
								if(k==0)
									pst.setInt(k+1, Integer.parseInt(cellValue));						
								else
									{pst.setString(k+1, cellValue);System.out.println(pst);}
								// 把一个SQL命令加入命令列表								
							} catch (NumberFormatException e) {
								// TODO Auto-generated catch block
								e.printStackTrace();
							} catch (SQLException e) {
								// TODO Auto-generated catch block
								e.printStackTrace();
							}
							
							sb.append(cellValue + "\t");
						}*/
						cellValue = cells[0].getContents();
						sb.append(cellValue + "\t");
						pst.setInt(1, Integer.parseInt(cellValue));	
						
						cellValue = cells[1].getContents();
						sb.append(cellValue + "\t");
						pst.setString(2, cellValue);
						
						pst.addBatch();
					}
					sb.append("\r\n");
				}
				sb.append("\r\n");
			}
		}
		 // 执行批量更新
	    pst.executeBatch();
	    // 语句执行完毕，提交本事务
	    con.commit();

	    pst.close();
	    con.close();
	    
		// 最后关闭资源，释放内存
		wb.close();
		return sb.toString();
	}

	/**
	 * 把內容寫入excel文件中
	 * 
	 * @param fileName
	 *            要寫入的文件的名稱
	 */
	public static void writeExcel(String fileName,String data) {
		WritableWorkbook wwb = null;
		try {
			// 首先要使用Workbook类的工厂方法创建一个可写入的工作薄(Workbook)对象
			wwb = Workbook.createWorkbook(new File(fileName));
		} catch (IOException e) {
			e.printStackTrace();
		}
		if (wwb != null) {
			// 创建一个可写入的工作表
			// Workbook的createSheet方法有两个参数，第一个是工作表的名称，第二个是工作表在工作薄中的位置
			WritableSheet ws = wwb.createSheet("sheet1", 0);
            
			// 下面开始添加单元格
			for (int i = 0; i < 10; i++) {
				for (int j = 0; j < 5; j++) {
					// 这里需要注意的是，在Excel中，第一个参数表示列，第二个表示行
					Label labelC = new Label(j, i, data);
					try {
						// 将生成的单元格添加到工作表中
						ws.addCell(labelC);
					} catch (RowsExceededException e) {
						e.printStackTrace();
					} catch (WriteException e) {
						e.printStackTrace();
					}

				}
			}

			try {
				// 从内存中写入文件中
				wwb.write();
				// 关闭资源，释放内存
				wwb.close();
			} catch (IOException e) {
				e.printStackTrace();
			} catch (WriteException e) {
				e.printStackTrace();
			}
		}
	}

}
