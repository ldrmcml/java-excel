package test;

import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.util.Region;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.File;

import java.sql.PreparedStatement;
import java.sql.ResultSetMetaData;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Connection;
import java.sql.ResultSet;
import java.io.UnsupportedEncodingException;

public class OperatorExcel2007 {
    private String fileSeparator = System.getProperty("file.separator");
    public OperatorExcel2007() {
    }

    public static void main(String agvs[]) {
		OperatorExcel2007 e = new OperatorExcel2007();
		// sql语句的要求：将数据表写在第一位如：select * from 数据表, 其他表…… where 数据表.id in (select id from 任意表)
		String[] sql = new String[]{"select * from table_a where a='a'",
			"select * from table_b where b='b'",
			""};
		try {
			for (int i = 0; i < sql.length; i++) {
			if (sql[i] != "") {
				e.createExcel(new String(sql[i].getBytes("GBK"), "8859_1"));//将sql中的字符串转换为“GBK”字节数组
			}
			}
		}
		catch (UnsupportedEncodingException ex) {
		}
    }

    /**
     * 获取标题单元格样式
     * @param wb HSSFWorkbook
     * @return HSSFCellStyle
     */
    private HSSFCellStyle getHeadCellStyle (HSSFWorkbook workBook) {
		HSSFCellStyle cellStyle = workBook.createCellStyle(); // 让HSSFWorkbook创建一个单元格样式的对象
		cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 设置单元格的横向对齐方式，具体参数参考HSSFCellStyle
		cellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER); // 设置单元格的纵向对齐方式，具体参数参考HSSFCellStyle
		cellStyle.setWrapText(true); // 设置单元格的文本方式为可多行编写方式
		/*
		// 设置单元格的填充方式，以及前景颜色和背景颜色
		// 三点注意：
		// 1.如果需要前景颜色或背景颜色，一定要指定填充方式，两者顺序无所谓；
		// 2.如果同时存在前景颜色和背景颜色，前景颜色的设置要写在前面；
		// 3.前景颜色不是字体颜色。
		cellStyle.setFillPattern(HSSFCellStyle.DIAMONDS); // 指定填充方式
		cellStyle.setFillForegroundColor(HSSFColor.RED.index); // 前景色
		cellStyle.setFillBackgroundColor(HSSFColor.LIGHT_YELLOW.index); // 背景色
		// 设置单元格底部的边框及其样式和颜色
		// 这里仅设置了底边边框，左边框、右边框和顶边框同理可设
		cellStyle.setBorderBottom(HSSFCellStyle.BORDER_SLANTED_DASH_DOT);
		cellStyle.setBottomBorderColor(HSSFColor.DARK_RED.index);
		*/
		HSSFFont font = workBook.createFont(); // 创建一个字体对象，因为字体也是单元格格式的一部分，所以从属于HSSFCellStyle
		font.setFontName("宋体"); // 设置字体
		// font.setItalic(true); // 斜体
		font.setColor(HSSFColor.BLUE.index);
		font.setFontHeightInPoints((short) 20); // 字体大小
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD); // 粗体字
		cellStyle.setFont(font); // 将字体对象赋值给单元格样式对象
		return cellStyle;
    }

    /**
     * 获取表头样式
     * @param wb HSSFWorkbook
     * @return HSSFCellStyle
     */
    private HSSFCellStyle getTitleCellStyle (HSSFWorkbook workBook) {
		HSSFCellStyle cellStyle = workBook.createCellStyle(); // 让HSSFWorkbook创建一个单元格样式的对象
		cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 设置单元格的横向对齐方式，具体参数参考HSSFCellStyle
		cellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER); // 设置单元格的纵向对齐方式，具体参数参考HSSFCellStyle
		cellStyle.setWrapText(true); // 设置单元格的文本方式为可多行编写方式
		HSSFFont font = workBook.createFont(); // 创建一个字体对象，因为字体也是单元格格式的一部分，所以从属于HSSFCellStyle
		font.setFontHeightInPoints((short) 12); // 字高
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD); // 粗体字
		cellStyle.setFont(font); // 将字体对象赋值给单元格样式对象
		return cellStyle;
    }

    /**
     * 获取默认文本类型单元格样式
     * @param wb HSSFWorkbook
     * @return HSSFCellStyle
     */
    private HSSFCellStyle getNormalTextCellStyle (HSSFWorkbook workBook) {
		HSSFCellStyle cellStyle = workBook.createCellStyle(); // 让HSSFWorkbook创建一个单元格样式的对象
		cellStyle.setAlignment(HSSFCellStyle.ALIGN_JUSTIFY); // 设置单元格的横向对齐方式，具体参数参考HSSFCellStyle
		cellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER); // 设置单元格的纵向对齐方式，具体参数参考HSSFCellStyle
		cellStyle.setWrapText(true); // 设置单元格的文本方式为可多行编写方式
		HSSFFont font = workBook.createFont(); // 创建一个字体对象，因为字体也是单元格格式的一部分，所以从属于HSSFCellStyle
		font.setFontHeightInPoints((short) 12); // 字高
		cellStyle.setFont(font); // 将字体对象赋值给单元格样式对象
		return cellStyle;
    }

    /**
     * 获取默认数值类型单元格样式
     * @param wb HSSFWorkbook
     * @return HSSFCellStyle
     */
    private HSSFCellStyle getNormalNumberCellStyle (HSSFWorkbook workBook) {
		HSSFCellStyle cellStyle = workBook.createCellStyle(); // 让HSSFWorkbook创建一个单元格样式的对象
		cellStyle.setAlignment(HSSFCellStyle.ALIGN_JUSTIFY); // 设置单元格的横向对齐方式，具体参数参考HSSFCellStyle
		cellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER); // 设置单元格的纵向对齐方式，具体参数参考HSSFCellStyle
		cellStyle.setWrapText(true); // 设置单元格的文本方式为可多行编写方式
		HSSFFont font = workBook.createFont(); // 创建一个字体对象，因为字体也是单元格格式的一部分，所以从属于HSSFCellStyle
		font.setFontHeightInPoints((short) 12); // 字高
		cellStyle.setFont(font); // 将字体对象赋值给单元格样式对象
		return cellStyle;
    }

    /**
     * 获取默认货币类型单元格样式
     * @param wb HSSFWorkbook
     * @return HSSFCellStyle
     */
    private HSSFCellStyle getNormalMoneyCellStyle (HSSFWorkbook workBook) {
		HSSFCellStyle cellStyle = workBook.createCellStyle(); // 让HSSFWorkbook创建一个单元格样式的对象
		cellStyle.setAlignment(HSSFCellStyle.ALIGN_JUSTIFY); // 设置单元格的横向对齐方式，具体参数参考HSSFCellStyle
		cellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER); // 设置单元格的纵向对齐方式，具体参数参考HSSFCellStyle
		cellStyle.setWrapText(true); // 设置单元格的文本方式为可多行编写方式
		HSSFFont font = workBook.createFont(); // 创建一个字体对象，因为字体也是单元格格式的一部分，所以从属于HSSFCellStyle
		font.setFontHeightInPoints((short) 12); // 字高
		cellStyle.setFont(font); // 将字体对象赋值给单元格样式对象
		return cellStyle;
    }

    /**
     * 设置表头文本
     * @param workBook HSSFWorkbook
     * @param st HSSFSheet
     */
    private void setHeadText (HSSFWorkbook workBook, HSSFSheet st, short colCount, String headText) {
		HSSFRow row = null; // 行
		HSSFCell cell = null; // 单元格
		row = st.createRow(0); // 创建属于上面Sheet的Row，参数0可以是0～65535之间的任何一个。行
		row.setHeightInPoints((float)30);
		cell = row.createCell((short) 0); // 创建属于上面Row的Cell，参数0可以是0～255之间的任何一个。列
		cell.setCellType(HSSFCell.CELL_TYPE_STRING); // 设置此单元格的格式为文本，此句可以省略，Excel会自动识别。
		cell.setCellValue(new HSSFRichTextString(headText)); // 此处是3.0.1版的改进之处，上一版可以直接setCellValue("Hello, World!")
		HSSFCellStyle cellStyle = getHeadCellStyle(workBook); // 将单元格样式对应应用于单元格
		cell.setCellStyle(cellStyle);
		st.addMergedRegion(new Region(0, (short) 0, 0, (short) colCount)); // 合并单元格：参数：起始行、列，结束行、列
    }

    public void createExcel(String sql) {
		HSSFWorkbook workBook = new HSSFWorkbook(); // 创建一个空白的WorkBook
		int pos = sql.indexOf("from ") + 5;
		String headText = sql.substring(pos, sql.indexOf(" ", pos)).trim();
		HSSFSheet st = workBook.createSheet(headText); // 基于上面的WorkBook创建属于此WorkBook的Sheet
		HSSFRow row = null; // 行
		HSSFCell cell = null; // 单元格
		short colCount = -1;
		Connection conn = null;
		PreparedStatement prestmt = null;
		ResultSet rs = null;
		try {
			conn = OperatorDataBase.getConnection();
			prestmt = conn.prepareStatement(sql);
			rs = prestmt.executeQuery();
			ResultSetMetaData rsmd = rs.getMetaData();
			colCount = (short)rsmd.getColumnCount();
			int i = 0;
			row = st.createRow(1); // 创建属于上面Sheet的Row，参数0可以是0～65535之间的任何一个。行
			setHeadText (workBook, st, colCount, headText);
			HSSFCellStyle titleCellStyle = getTitleCellStyle(workBook);
			HSSFCellStyle textCellStyle = getNormalTextCellStyle(workBook); // 将单元格样式对应应用于单元格
			for (int j = 0; j < colCount; j++) { // 写列名
				cell = row.createCell(j); // 创建属于上面Row的Cell，参数0可以是0～255之间的任何一个。列
				cell.setCellType(HSSFCell.CELL_TYPE_STRING); // 设置此单元格的格式为文本，此句可以省略，Excel会自动识别。
				cell.setCellValue(new HSSFRichTextString(new String(rsmd.getColumnName(j + 1).getBytes("ISO8859_1"), "GBK"))); // 设置此单元格的
				cell.setCellStyle(titleCellStyle);
			}
			while (rs.next()) { // 写数据
				row = st.createRow(i + 2); // 创建属于上面Sheet的Row，参数0可以是0～65535之间的任何一个。行
				String value[] = new String[colCount];
				for (int j = 0; j < colCount; j++) {
					cell = row.createCell(j); // 创建属于上面Row的Cell，参数0可以是0～255之间的任何一个。列
					cell.setCellType(HSSFCell.CELL_TYPE_STRING); // 设置此单元格的格式为文本，此句可以省略，Excel会自动识别。
					cell.setCellValue(new HSSFRichTextString(new String((rs.getString(j + 1) == null ? "" :
					rs.getString(j + 1)).getBytes("8859_1"), "GBK"))); // 此处是3.0.1版的改进之处，上一版可以直接setCellValue("Hello, World!")
					cell.setCellStyle(textCellStyle);
				}
				i++;
			}
		}
		catch (SQLException e) {
			e.printStackTrace();
		}
		catch (ClassNotFoundException e) {
			e.printStackTrace();
		}
		catch (UnsupportedEncodingException e) {
			e.printStackTrace();
		}
		finally {
			try {
				OperatorDataBase.close(conn, prestmt, rs);
			}
			catch (SQLException e) {
				e.printStackTrace();
			}
		}
		st.addMergedRegion(new Region(0, (short) 0, 0, (short) colCount)); // 合并单元格：参数：起始行、列，结束行、列
		// setHeadText(workBook, st); // 设置头部分
		FileOutputStream outStream = null;
		File fExcel = null;
		try {
			fExcel = new File("C:" + fileSeparator + headText + ".xls");
			if (fExcel.exists()) {
				try {
					fExcel.delete();
				}
				catch (Exception e) {
					e.printStackTrace();
				}
			}
			outStream = new FileOutputStream(fExcel); // 创建一个文件输出流，指定文件保存目录。xls是Excel97-2003的标准扩展名，2007是xlsx，目前的POI能直接生产的还是xls格式
			workBook.write(outStream); // 把WorkBook写到流里
		}
		catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		catch (IOException e) {
			e.printStackTrace();
		}
		finally {
			try {
				outStream.close(); // 手动关闭流。官方文档已经做了特别说明，说POI不负责关闭用户打开的流。
			}
			catch (IOException e) {
				e.printStackTrace();
			}
		}
    }
}

class OperatorDataBase {

    /**
     * @title : 获取一个连接
     * @return Connection
     * @throws ClassNotFoundException
     * @throws SQLException
     */
    public static Connection getConnection() throws ClassNotFoundException, SQLException {
		Connection conn = null;
		conn = getConnection(true);
		return conn;
    }

    /**
     * @title : 获取一个连接
     * @param isAuto boolean
     * @return Connection
     * @throws ClassNotFoundException
     * @throws SQLException
     */
    public static Connection getConnection(boolean isAuto) throws ClassNotFoundException, SQLException {
		Connection conn = null;
		// 加载数据库驱动类
		Class.forName("oracle.jdbc.driver.OracleDriver");
		// jdbc:oracle:thin:@" + DB_URL + ":" + DB_Post + ":" + DB_Name
		conn = DriverManager.getConnection("jdbc:oracle:thin:@192.168.1.1:1521:ctais", "ctais2", "oracle");
		conn.setAutoCommit(isAuto);
		return conn;
    }

    /**
     * @title : 提交数据库链接
     * @param conn Connection
     * @throws SQLException
     */
    public static void commit(Connection conn) throws SQLException {
		if (conn != null && !conn.isClosed()) {
			conn.commit();
		}
    }

    /**
     * @title : 回滚数据库链接
     * @param conn Connection
     * @throws SQLException
     */
    public static void rollback(Connection conn) throws SQLException {
		if (conn != null && !conn.isClosed()) {
			conn.rollback();
		}
    }

    /**
     * @title : 关闭数据库连接
     * @param conn Connection
     * @throws SQLException
     */
    public static void close(Connection conn) throws SQLException {
		if (conn != null && !conn.isClosed()) {
			conn.close();
		}
		conn = null;
    }

    /**
     * @title : 关闭sql分析执行器
     * @param prestmt PreparedStatement
     * @throws SQLException
     */
    public static void close(PreparedStatement prestmt) throws SQLException {
		if (prestmt != null) {
			prestmt.close();
			prestmt = null;
		}
    }

    /**
     * @title : 关闭记录集
     * @param rest ResultSet
     * @throws SQLException
     */
    public static void close(ResultSet rest) throws SQLException {
		if (rest != null) {
			rest.close();
			rest = null;
		}
    }

    /**
     * @title : 关闭数据库连接、sql分析执行器、记录集
     * @param conn Connection
     * @param prestmt PreparedStatement
     * @param rest ResultSet
     * @throws SQLException
     */
    public static void close(Connection conn, PreparedStatement prestmt, ResultSet rest) throws SQLException {
		close(rest);
		close(prestmt);
		close(conn);
    }
}
