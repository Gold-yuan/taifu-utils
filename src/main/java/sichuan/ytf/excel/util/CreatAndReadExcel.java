package sichuan.ytf.excel.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreatAndReadExcel {

	public static void main(String[] args) throws Exception {
		/*
		String fileName = "abc.xlsx";
		String filePath = creat2003And2007Excel("abc.xlsx");
		// 读取2003Excel文件
		String path2003 = System.getProperty("user.dir") + System.getProperty("file.separator") + fileName;// 获取项目文件路径
		// +2003版文件名
		System.out.println("路径：" + path2003);
		File f2003 = new File(path2003);
		try {
			read2003And2007Excel(f2003);
			File file = new File(filePath);
			file.delete();
		} catch (IOException e) {
			e.printStackTrace();
		}
		*/
		
		String excelName = "test.xlsx";
		String[] sheetNames= {"sheet1","sheet22"};
		List<String[]> titleNames = new ArrayList<String[]>();
		List<String[]> fields = new ArrayList<String[]>();
		List<List<Map<String, Object>>> datas = new ArrayList<List<Map<String, Object>>>();
		
		String[] titleName0 = {"名字","年龄"};
		String[] titleName1 = {"性别","爱好"};
		titleNames.add(titleName0);
		titleNames.add(titleName1);
		String[] fields0 = {"name","age"};
		String[] fields1 = {"sex","f"};
		fields.add(fields0);
		fields.add(fields1);
		
		Map<String, Object> map1 = new HashMap<String, Object>();
		map1.put("name", "二狗子");
		map1.put("age", "22");
		Map<String, Object> map2 = new HashMap<String, Object>();
		map2.put("sex", "女");
		map2.put("f", "男");
		
		List<Map<String, Object>> sheet0 = new ArrayList<Map<String, Object>>();
		sheet0.add(map1);
		List<Map<String, Object>> sheet1 = new ArrayList<Map<String, Object>>();
		sheet1.add(map2);
		
		datas.add(sheet0);
		datas.add(sheet1);
		
		String path = creat2003And2007ExcelMultipartSheetByFileName(excelName, sheetNames, titleNames, fields, datas);
		System.out.println(path);
		File excel = new File(path);
		read2003And2007Excel(excel);
			
	}

	/**
	 * 创建excel
	 * 
	 * @param excelName
	 * @param sheetNames
	 * @param titleNames
	 * @param fields
	 * @param datas map-->row, list<map>-->sheet, list<list<map>>-->excel
	 * @return excel FilePath
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	public static String creat2003And2007ExcelMultipartSheetByFileName(
			String excelName, String[] sheetNames, List<String[]> titleNames,
			List<String[]> fields, List<List<Map<String, Object>>> datas)
			throws FileNotFoundException, IOException {
		if (datas == null)
			return null;
		
		Workbook workBook;
		boolean is03Excel = ".xls".endsWith(excelName);
		if (is03Excel) {
			workBook = new HSSFWorkbook();// 创建 一个2003 excel文档对象
		} else {
//			workBook = new XSSFWorkbook();// 创建 一个2007 excel文档对象
			workBook = new SXSSFWorkbook(1000);// //内存中保留 10000 条数据，以免内存溢出，其余写入 硬盘        
		}

		// 样式
		CellStyle contentStyle = workBook.createCellStyle();// 创建样式对象
		contentStyle.setBorderTop(BorderStyle.THIN);// 设置四周边框边框
		contentStyle.setBorderBottom(BorderStyle.THIN);
		contentStyle.setBorderLeft(BorderStyle.THIN);
		contentStyle.setBorderRight(BorderStyle.THIN);
		
		// 表头样式
		CellStyle titleStyle = workBook.createCellStyle();// 创建样式对象
		titleStyle.setBorderTop(BorderStyle.THIN);// 设置四周边框边框
		titleStyle.setBorderBottom(BorderStyle.THIN);
		titleStyle.setBorderLeft(BorderStyle.THIN);
		titleStyle.setBorderRight(BorderStyle.THIN);
//		titleStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		titleStyle.setFillBackgroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		

		// 数据填充
		for (int i = 0; i < sheetNames.length; i++) {
			Sheet sheet = workBook.createSheet(sheetNames[i]);// 创建一个工作薄对象

			Row rowTitle = sheet.createRow(0);// 创建一个行对象 0开始
			for (int k = 0; k < titleNames.get(i).length; k++) {
				sheet.setColumnWidth(k, 5000);// 设置第二列的宽度为4000
				String name = titleNames.get(i)[k];
				Cell cell = rowTitle.createCell(k);// 创建单元格0开始
				cell.setCellValue(name);
				cell.setCellStyle(titleStyle);// 应用样式对象
			}

			if (datas.size() > i) {
				List<Map<String, Object>> sheetRows = datas.get(i);
				ZipSecureFile.setMinInflateRatio(0l);  
				for (int j = 0; j < sheetRows.size(); j++) {
					if(j>1000 && (j+1)%1000==0) {
						System.out.println((j + "/" + sheetRows.size()) + ", "+ j*100/sheetRows.size()+"%");
					}
					HashMap<String, Object> dataObj = (HashMap<String, Object>) sheetRows.get(j);
					Row rowContent = sheet.createRow(j+1);// 创建一个行对象j+1是第二行，第一行是表头
					for (int x = 0; x < fields.get(i).length; x++) {
						String field = fields.get(i)[x];
						Cell cell = rowContent.createCell(x);// 创建单元格0开始
						Object val = dataObj.get(field);
						String value = val==null?"":val.toString();
						cell.setCellValue(value);
						cell.setCellStyle(contentStyle);// 应用样式对象
					}
					rowContent = null;
				}
			}
		}
		datas = null;
		// 文件输出流
		File file = new File(excelName);
		FileOutputStream os = new FileOutputStream(file);
		workBook.write(os);// 将文档对象写入文件输出流
		os.close();// 关闭文件输出流
		workBook.close();
		return file.getAbsolutePath();
	}

	/**
	 * 创建2003 2007版Excel文件
	 * @return 
	 * 
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	public static String creat2003And2007Excel(String excelName)
			throws FileNotFoundException, IOException {
		Workbook workBook;
		boolean is03Excel = ".xls".endsWith(excelName);
		if (is03Excel) {
			workBook = new HSSFWorkbook();// 创建 一个2003 excel文档对象
		} else {
			workBook = new XSSFWorkbook();// 创建 一个2007 excel文档对象
		}
		Sheet sheet = workBook.createSheet();// 创建一个工作薄对象

		sheet.setColumnWidth(1, 10000);// 设置第二列的宽度为

		Row row = sheet.createRow(1);// 创建一个行对象

		row.setHeightInPoints(23);// 设置行高23像素

		CellStyle style = workBook.createCellStyle();// 创建样式对象

		// 设置字体

		Font font = workBook.createFont();// 创建字体对象

		font.setFontHeightInPoints((short) 15);// 设置字体大小

		font.setBold(true);

		font.setFontName("黑体");// 设置为黑体字

		style.setFont(font);// 将字体加入到样式对象

		// 设置对齐方式

		style.setAlignment(HorizontalAlignment.CENTER);

		style.setVerticalAlignment(VerticalAlignment.CENTER);// 垂直居中

		// 设置边框
		style.setBorderTop(BorderStyle.THICK);// 顶部边框粗线
		style.setTopBorderColor(HSSFColor.RED.index);// 设置为红色
		style.setBorderBottom(BorderStyle.DOUBLE);// 底部边框双线
		style.setBorderLeft(BorderStyle.THIN);// 左边边框
		style.setBorderRight(BorderStyle.THIN);// 右边边框
		style.setDataFormat(HSSFDataFormat.getBuiltinFormat("m/d/yy h:mm"));// 格式化日期

		Cell cell = row.createCell(1);// 创建单元格
		cell.setCellValue(new Date());// 写入当前日期
		cell.setCellStyle(style);// 应用样式对象

		// 文件输出流
		File file = new File(excelName);
		FileOutputStream os = new FileOutputStream(file);
		workBook.write(os);// 将文档对象写入文件输出流

		os.close();// 关闭文件输出流
		workBook.close();
		return file.getAbsolutePath();
	}

	@SuppressWarnings("deprecation")
	private static List<List<Object>> read2003And2007Excel(File file)
			throws IOException {

		List<List<Object>> list = new LinkedList<List<Object>>();

		Workbook xwb;
		if (file.getName().endsWith(".xls")) {
			xwb = new HSSFWorkbook(new FileInputStream(file));// 创建一个2003excel文档对象
			System.out.println("读取office 2003 excel内容如下：");
		} else {
			xwb = new XSSFWorkbook(new FileInputStream(file));// 创建一个2007excel文档对象
			System.out.println("读取office 2007 excel内容如下：");
		}
		for (int z = 0; z < xwb.getNumberOfSheets(); z++) {
			// 读取第一章表格内容
			Sheet sheet = xwb.getSheetAt(z);
			Object value = null;
			Row row = null;
			Cell cell = null;
			for (int i = sheet.getFirstRowNum(); i <= sheet
					.getPhysicalNumberOfRows(); i++) {
				row = sheet.getRow(i);
				if (row == null) {
					continue;
				}
				List<Object> linked = new LinkedList<Object>();
				for (int j = row.getFirstCellNum(); j <= row.getLastCellNum(); j++) {
					cell = row.getCell(j);
					if (cell == null) {
						continue;
					}
					DecimalFormat df = new DecimalFormat("0");// 格式化 number
																// String
					// 字符
					SimpleDateFormat sdf = new SimpleDateFormat(
							"yyyy-MM-dd HH:mm:ss");// 格式化日期字符串
					DecimalFormat nf = new DecimalFormat("0.00");// 格式化数字

					switch (cell.getCellType()) {
					case Cell.CELL_TYPE_STRING:
						// System.out.println(i + "行" + j +
						// " 列 is String type");
						value = cell.getStringCellValue();
						System.out.print("  " + value + "  ");
						break;
					case Cell.CELL_TYPE_NUMERIC:
						// System.out.println(i + "行" + j
						// + " 列 is Number type ; DateFormt:"
						// + cell.getCellStyle().getDataFormatString());
						if ("@".equals(cell.getCellStyle()
								.getDataFormatString())) {
							value = df.format(cell.getNumericCellValue());

						} else if ("General".equals(cell.getCellStyle()
								.getDataFormatString())) {
							value = nf.format(cell.getNumericCellValue());
						} else {
							value = sdf.format(HSSFDateUtil.getJavaDate(cell
									.getNumericCellValue()));
						}
						System.out.print("  " + value + "  ");
						break;
					case Cell.CELL_TYPE_BOOLEAN:
						// System.out.println(i + "行" + j +
						// " 列 is Boolean type");
						value = cell.getBooleanCellValue();
						System.out.print("  " + value + "  ");
						break;
					case Cell.CELL_TYPE_BLANK:
						// System.out.println(i + "行" + j + " 列 is Blank type");
						value = "";
						// System.out.println(value);
						break;
					default:
						// System.out.println(i + "行" + j +
						// " 列 is default type");
						value = cell.toString();
						System.out.print("  " + value + "  ");
					}
					if (value == null || "".equals(value)) {
						continue;
					}
					linked.add(value);
				}
				System.out.println("");
				list.add(linked);
			}
		}
		xwb.close();
		return list;
	}
}