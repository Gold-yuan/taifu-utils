package sichuan.ytf.excel.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * 
 * 基于POI的读写Excel文件的工具类
 * 
 */
public final class ExcelUtils {

    public static Map<String, List<Map<String, Object>>> excelMap;

    public static void main(String[] args) throws Exception {
//        List<Map<String, Object>> readExcel = ExcelUtils.readExcelJob(new FileInputStream(
//                "E:\\develop\\workspace\\eclipse\\TrustedSDP-data\\src\\main\\resources\\templates\\java\\job\\食药监网站-数据爬取整理.xlsx"),
//                "北京");
//        System.out.println(readExcel);
        System.out.println(loadExcel());
    }
    

    public static Map<String, List<Map<String, Object>>> loadExcel() throws FileNotFoundException, Exception {
        if (excelMap == null) {
            String eclipseWorkspaceDir = System.getProperty("user.dir") + File.separator;
            // excel路径
            String filePath = eclipseWorkspaceDir + "src/main/resources/templates/java/other/食药监网站-数据爬取整理.xlsx";
            excelMap = ExcelUtils.readExcelJob(new FileInputStream(filePath));
        }
        return excelMap;
    }

    public static Map<String, List<Map<String, Object>>> loadExcel(InputStream input)
            throws FileNotFoundException, Exception {
        if (excelMap == null) {
            excelMap = ExcelUtils.readExcelJob(input);
        }
        return excelMap;
    }

    private static <T> Field[] getFields(Class<T> clazz, String[] fieldNames) {
        Field[] fs = new Field[fieldNames.length];
        Field[] fields = clazz.getDeclaredFields();
        for (int i = 0; i < fieldNames.length; i++) {
            O: for (Field field : fields) {
                field.setAccessible(true);
                if (field.getName().equals(fieldNames[i])) {
                    fs[i] = field;
                    break O;
                }
            }
        }
        System.out.println(Arrays.toString(fs));
        return fs;
    }

    private static <T> String[][] getDatas(String[] titles, List<T> beans, String[] fieldNames) throws Exception {
        List<String[]> datas = new ArrayList<>();

        if (titles != null && titles.length > 0) {
            datas.add(titles);
        }
        // 根据List集合中的JavaBean对象的类型，和参数fieldNames，获取要写出的字段数组；
        Field[] fs = getFields(beans.get(0).getClass(), fieldNames);
        for (T t : beans) {
            String[] data = new String[fs.length];
            for (int i = 0; i < fs.length; i++) {
                if (fs[i] != null) {
                    Object obj = fs[i].get(t);
                    String value = obj == null ? "" : obj.toString();
                    data[i] = value;
                }
            }
            datas.add(data);
        }
        return datas.toArray(new String[datas.size()][]);
    }

    /**
     * 将指定的二维字符串数组中，写出到指定的Excel文件中
     * 
     * @param datas：保存了要写出的二维数组；
     * @param output：关联到要输出的Excel文件的输出流
     * @throws IOException
     * @throws FileNotFoundException
     */
    public static void writeToExcel(String[][] datas, OutputStream output) throws IOException, FileNotFoundException {
        // 创建一个Workbook对象；
        Workbook book = new HSSFWorkbook();
        // 通过workbook对象，创建一个工作表（sheet对象）
        Sheet sheet = book.createSheet();
        for (int i = 0; i < datas.length; i++) {
            // 通过工作表创建行
            Row row = sheet.createRow(i);
            for (int j = 0; j < datas[i].length; j++) {
                // 通过行创建单元格
                Cell cell = row.createCell(j);
                cell.setCellType(CellType.STRING);
                // 为单元格设置内容
                cell.setCellValue(datas[i][j]);
            }
        }
        book.write(output);
        book.close();
    }

    /**
     * 将指定的二维字符串数组中，写出到指定的Excel文件中
     * 
     * @param datas：保存了要写出的二维数组；
     * @param output：关联到要输出的Excel文件的输出流
     * @throws IOException
     * @throws FileNotFoundException
     */
    public static void writeToExcelStyle(String[][] datas, OutputStream output)
            throws IOException, FileNotFoundException {
        // 创建一个Workbook对象；
        Workbook book = new HSSFWorkbook();
        // 通过workbook对象，创建一个工作表（sheet对象）
        Sheet sheet = book.createSheet();
        // 设置列宽100px
        sheet.setColumnWidth(0, (int) 35.7 * 200);
        for (int i = 0; i < datas.length; i++) {
            // 通过工作表创建行
            Row row = sheet.createRow(i);
            // 设置行高40px
            row.setHeight((short) (15.625 * 40));
            row.setHeightInPoints((float) 40);

            // 生成一个样式
            CellStyle style = book.createCellStyle();
            // 设置这些样式
            style.setAlignment(HorizontalAlignment.CENTER);// 水平居中
            style.setVerticalAlignment(VerticalAlignment.CENTER);// 垂直居中
            style.setBorderBottom(BorderStyle.THIN);
            style.setBorderLeft(BorderStyle.THIN);
            style.setBorderRight(BorderStyle.THIN);
            style.setBorderTop(BorderStyle.THIN);
            for (int j = 0; j < datas[i].length; j++) {
                // 通过行创建单元格
                Cell cell = row.createCell(j);

                cell.setCellStyle(style);
                cell.setCellType(CellType.STRING);
                // 为单元格设置内容
                cell.setCellValue(datas[i][j]);
            }
        }
        book.write(output);
        book.close();
    }

    /**
     * 将指定集合中的所有JavaBean数据，写出到关联指定Excel文件的输出流中；
     * 
     * @param titles：输出的Excel文件的标题行的内容
     * @param fieldNames：一行中的各个列对应的JavaBean中的字段
     * @param beans：要输出的所有JavaBean对象
     * @param output：关联到指定Excel文件的输出流
     * @throws Exception
     */
    public static <T> void writeToExcel(String[] titles, String[] fieldNames, List<T> beans, OutputStream output)
            throws Exception {
        String[][] datas = getDatas(titles, beans, fieldNames);
        writeToExcel(datas, output);
    }

    public static List<Map<String, Object>> readExcelJob(InputStream input, String area) throws Exception {
        List<Map<String, Object>> list = new ArrayList<>();
        // 创建一个Workbook对象
        Workbook book = WorkbookFactory.create(input);
        for (int sheetIndex = 0; sheetIndex < book.getNumberOfSheets(); sheetIndex++) {
            // 通过Workbook对象，获取里面的工作表（sheet对象）
            Sheet sheet = book.getSheetAt(sheetIndex);
            String sheetName = sheet.getSheetName();
            if (!sheetName.equals(area)) {
                continue;
            }

            // 遍历第二行的所有值，作为map的key
            List<String> mapKey = new ArrayList<>();
            Row rowField = sheet.getRow(1);
            for (int j = 0; j < rowField.getLastCellNum(); j++) {
                Cell cell = rowField.getCell(j);
                cell.setCellType(CellType.STRING);
                mapKey.add(cell.toString());
            }

            // 遍历第三行起的所有制，作为map的value，每行对应一个map，即
            for (int i = 2; i <= sheet.getLastRowNum(); i++) {
                Map<String, Object> map = new HashMap<>();
                // 通过行号获取行
                Row row = sheet.getRow(i);
                // 遍历获取一行中的所有单元格
                for (int j = 0; j < row.getLastCellNum(); j++) {
                    Cell cell = row.getCell(j);
                    System.out.print("第" + i + "行" + j + "列，");
                    if (cell == null) {
                        map.put(mapKey.get(j), "");
                    } else {
                        cell.setCellType(CellType.STRING);
                        map.put(mapKey.get(j), cell.toString());
                    }
                }
                list.add(map);
            }
        }
        return list;
    }

    /**
     * @param input
     * @return Map<sheetName, List<Map<第二行值，其他行值>> <br>
     *         Excel<sheetName sheetVal<row<cellkey,cellval>>
     * @throws Exception
     */
    public static Map<String, List<Map<String, Object>>> readExcelJob(InputStream input) throws Exception {
        Map<String, List<Map<String, Object>>> excel = new HashMap<String, List<Map<String, Object>>>();
        // 创建一个Workbook对象
        Workbook book = WorkbookFactory.create(input);
        for (int sheetIndex = 0; sheetIndex < book.getNumberOfSheets(); sheetIndex++) {
            // 通过Workbook对象，获取里面的工作表（sheet对象）
            Sheet sheet = book.getSheetAt(sheetIndex);

            // 遍历第二行的所有值，作为map的key
            List<String> mapKey = new ArrayList<>();
            Row rowField = sheet.getRow(1);
            for (int j = 0; j < rowField.getLastCellNum(); j++) {
                Cell cell = rowField.getCell(j);
                cell.setCellType(CellType.STRING);
                mapKey.add(cell.toString());
            }

            // 遍历第三行起的所有制，作为map的value，每行对应一个map，即
            List<Map<String, Object>> list = new ArrayList<>();
            System.out.println("\n"+sheet.getSheetName());
            for (int i = 2; i <= sheet.getLastRowNum(); i++) {
                Map<String, Object> map = new HashMap<>();
                // 通过行号获取行
                Row row = sheet.getRow(i);
                // 遍历获取一行中的所有单元格
                System.out.print("第" + i + "行");
                for (int j = 0; j < row.getLastCellNum(); j++) {
                    Cell cell = row.getCell(j);
                    if (cell == null) {
                        map.put(mapKey.get(j), "");
                    } else {
                        cell.setCellType(CellType.STRING);
                        map.put(mapKey.get(j), cell.toString());
                    }
                }
                list.add(map);
            }
            String sheetName = sheet.getSheetName();
            excel.put(sheetName, list);
        }
        return excel;
    }

    /**
     * 从指定的Excel的输入流中读取数据，写到指定类型的bean对象的List集合中；
     * 
     * @param input：关联到某个Excel文件的输入流
     * @param clazz：要封装一行数据的JavaBean的类型
     * @param fieldNames：一行数据中的列按顺序和JavaBean中对应的字段
     * @return
     * @throws Exception
     */
    public static <T> List<T> readExcel(InputStream input, Class<T> clazz, String[] fieldNames) throws Exception {
        List<T> list = new ArrayList<>();
        Field[] fs = getFields(clazz, fieldNames);
        // 创建一个Workbook对象
        Workbook book = WorkbookFactory.create(input);
        // 通过Workbook对象，获取里面的工作表（sheet对象）
        Sheet sheet = book.getSheetAt(0);
        // 获取工作表中的行
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            // 通过行号获取行
            Row row = sheet.getRow(i);
            T t = clazz.newInstance();
            // 遍历获取一行中的所有单元格
            for (int j = 0; j < fs.length; j++) {
                if (fs[j] != null) {
                    Cell cell = row.getCell(j);
                    cell.setCellType(CellType.STRING);
                    setCellValue(fs[j], t, cell);
                }
            }
            list.add(t);
        }
        return list;
    }

    private static void setCellValue(Field field, Object bean, Cell cell) throws Exception {
        String str = cell.toString();
        @SuppressWarnings("rawtypes")
        Class clazz = field.getType();
        if (clazz == byte.class) {
            field.set(bean, Byte.parseByte(str));
        } else if (clazz == short.class) {
            field.set(bean, Short.parseShort(str));
        } else if (clazz == int.class) {
            field.set(bean, Integer.parseInt(str));
        } else if (clazz == long.class) {
            field.set(bean, Long.parseLong(str));
        } else if (clazz == double.class) {
            field.set(bean, Double.parseDouble(str));
        } else if (clazz == float.class) {
            field.set(bean, Float.parseFloat(str));
        } else if (clazz == boolean.class) {
            field.set(bean, Boolean.parseBoolean(str));
        } else {
            field.set(bean, str);
        }
    }
}