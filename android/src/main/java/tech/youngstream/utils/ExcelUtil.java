package tech.youngstream.utils;
import android.text.TextUtils;

import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFName;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

public class ExcelUtil {
    private static final DecimalFormat df = new DecimalFormat("0");// 格式化 number String 字符
    private static final SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");// 格式化日期字符串
    private static final DecimalFormat nf = new DecimalFormat("0.00");// 格式化数字

    /**
     * 获取某单元格的值
     * @param cell
     * @return
     */
    public static Object getCellValue(Cell cell) {
        Object object = null;
        if(cell==null) {
            return "";
        }
        switch (cell.getCellType()) {
            case _NONE:
                break;
            case NUMERIC: // 数字
                object = cell.getNumericCellValue();
                break;
            case STRING: // 字符串
                object = cell.getStringCellValue();
                break;
            case BOOLEAN: // Boolean
                object = cell.getBooleanCellValue();
                break;
            case FORMULA: // 公式
                object = cell.getCellFormula();
                break;
            case BLANK: // 空值
                object = null;
                break;
            case ERROR: // 故障
                System.out.print(" ");
                break;
            default:
                System.out.print("未知类型   ");
                break;
        }
        if (CellType.NUMERIC == cell.getCellType()) {
            //判断是否为日期类型
            if (HSSFDateUtil.isCellDateFormatted(cell)) {
                //用于转化为日期格式
                Date d = cell.getDateCellValue();
                object = sdf.format(d);
            } else if("@".equals(cell.getCellStyle().getDataFormatString())){
                object = df.format(cell.getNumericCellValue());
            } else if("General".equals(cell.getCellStyle().getDataFormatString())){
                object = nf.format(cell.getNumericCellValue());
            }
        }
        if(object == null || TextUtils.isEmpty(object.toString())) {
            object = "";
        }
        return object;
    }
    /**
     * 获取excel第column列的表头
     * @param sheet
     * @param column
     * @return
     */
    public static String getCellHead(Sheet sheet, int column) {
        Cell cell = sheet.getRow(0).getCell(column);
        return cell.getStringCellValue();
    }
    /**
     * 根据标题和列内容生产表
     * @param book
     * @param title 标题
     * @param columnNames 各列标题
     * @return 表第0行是表的标题，第1行是表的列的标题
     */
    public static XSSFSheet createSheetByTitles(XSSFWorkbook book, String title, String[] columnNames) {
        XSSFSheet sheet = book.createSheet();
        XSSFRow titleRow = sheet.createRow(0);
        titleRow.setHeight((short) 500);
        //生成字体
        XSSFFont font = book.createFont();
        font.setFontName("微软雅黑");
        font.setFontHeight((short) 300);
        // 创建单元格样式
        XSSFCellStyle style = book.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        //设置字体
        style.setFont(font);

        XSSFCell titleCell = titleRow.createCell(0);
        //合并单元格(startRow，endRow，startColumn，endColumn)
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, columnNames.length-1));
        //设置表的标题样式
        titleCell.setCellStyle(style);
        titleCell.setCellValue(title);

        XSSFRow columnTitleRow = sheet.createRow(1);

        XSSFFont columnTitleFont = book.createFont();
        columnTitleFont.setFontName("微软雅黑");

        XSSFCellStyle columnTitleStyle = book.createCellStyle();
        columnTitleStyle.setAlignment(HorizontalAlignment.CENTER);
        columnTitleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        columnTitleStyle.setFont(columnTitleFont);

        for (int i = 0; i < columnNames.length; i++) {
            XSSFCell iCell = columnTitleRow.createCell(i);
            iCell.setCellStyle(columnTitleStyle);
            iCell.setCellValue(columnNames[i]);
            iCell.setCellType(CellType.STRING);
            sheet.setColumnWidth(i, 6000);
        }
        return sheet;
    }

    public static XSSFSheet createSheetByTitles(XSSFWorkbook book, String[] columnNames) {
        XSSFSheet sheet = book.createSheet();
        //生成字体
        XSSFFont font = book.createFont();
        font.setFontName("微软雅黑");
        font.setFontHeight((short) 300);
        // 创建单元格样式
        XSSFCellStyle style = book.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        //设置字体
        style.setFont(font);

        XSSFRow columnTitleRow = sheet.createRow(0);

        XSSFFont columnTitleFont = book.createFont();
        columnTitleFont.setFontName("微软雅黑");

        XSSFCellStyle columnTitleStyle = book.createCellStyle();
        columnTitleStyle.setAlignment(HorizontalAlignment.CENTER);
        columnTitleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        columnTitleStyle.setFont(columnTitleFont);
        for (int i = 0; i < columnNames.length; i++) {
            XSSFCell iCell = columnTitleRow.createCell(i);
            iCell.setCellStyle(columnTitleStyle);
            iCell.setCellValue(columnNames[i]);
            iCell.setCellType(CellType.STRING);
            sheet.setColumnWidth(i, 6000);
        }
        return sheet;
    }
    /**
     * 生成表格下拉列表约束
     * @param row 行
     * @param col 列
     * @param values 下拉列表值[]
     * @return
     */
    public static DataValidation createSelectValidation(XSSFSheet sheet,int row,int col,String[] values) {
        //(int firstRow, int lastRow,int firstCol,int lastCol)
        CellRangeAddressList regions = new CellRangeAddressList(row, 100000, col, col);

        DataValidationHelper helper = sheet.getDataValidationHelper();

        //CellRangeAddressList(firstRow, lastRow, firstCol, lastCol)设置行列范围
        CellRangeAddressList addressList = new CellRangeAddressList(row, Integer.MAX_VALUE, col, col);

        //设置下拉框数据
        DataValidationConstraint constraint = helper.createExplicitListConstraint(values);
        DataValidation dataValidation = helper.createValidation(constraint, addressList);
        //处理Excel兼容性问题
        if(dataValidation instanceof XSSFDataValidation) {
            dataValidation.setSuppressDropDownArrow(true);
            dataValidation.setShowErrorBox(true);
        }else {
            dataValidation.setSuppressDropDownArrow(false);
        }
        // 绑定下拉框和作用区域
        return dataValidation;
    }

    /**
     * 获取excel表头 均从第一行开始
     * @param sheet
     * @return
     */
    public static String[] getSheetHead(Sheet sheet,int startRow) {
        Row titleRow = sheet.getRow(startRow);
        List<String> titles = new ArrayList<String>();
        if(titleRow == null) {
            System.out.print("该excel无任何数据");
            return null;
        }
        for(int i = titleRow.getFirstCellNum(); i < titleRow.getLastCellNum(); i++) {
            titles.add(getCellValue(titleRow.getCell(i)).toString());
        }
        return titles.toArray(new String[]{});
    }

    /**
     * 判断该excel文件的表头是否为数组headers中的值相同
     * @param file
     * @param headers
     * @return
     */
    public static boolean isValidFile(File file, String[] headers, int rowNo) {
        Workbook book = null;
        List<String> headerList = Arrays.asList(headers);
        try{
            if(file == null ) {
                return false;
            }
            if(file.getName().endsWith(".xls")) {
                book = new HSSFWorkbook(new FileInputStream(file));
            } else if(file.getName().endsWith(".xlsx")) {
                book = new XSSFWorkbook(new FileInputStream(file));
            }
            if(book == null) return false;
            Sheet sheet = book.getSheetAt(0);
            if(sheet==null) {
                System.out.print("ERROR: 文档0获取失败");
                return false;
            }
            List<String> fileHeaders = Arrays.asList(getSheetHead(sheet,rowNo));
            System.out.print("std header : "+headerList);
            System.out.print("tar header : "+fileHeaders);
            return fileHeaders.containsAll(headerList);
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        } finally {
            try {
                if(book != null)
                    book.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 创建名称
     * @param wb
     * @param name
     * @param expression
     * @return
     */
    public static HSSFName createName(HSSFWorkbook wb, String name, String expression){
        HSSFName refer = wb.createName();
        refer.setRefersToFormula(expression);
        refer.setNameName(name);
        return refer;
    }
    /**
     * 设置数据有效性（通过名称管理器级联相关）
     * @param name
     * @param firstRow
     * @param endRow
     * @param firstCol
     * @param endCol
     * @return
     */
    public static HSSFDataValidation setDataValidation(String name, int firstRow, int endRow, int firstCol, int endCol){
        //设置下拉列表的内容
        //加载下拉列表内容
        DVConstraint constraint = DVConstraint.createFormulaListConstraint(name);
        // 设置数据有效性加载在哪个单元格上。
        // 四个参数分别是：起始行、终止行、起始列、终止列
        CellRangeAddressList regions = new CellRangeAddressList(firstRow, endRow, firstCol, endCol);
        // 数据有效性对象
        HSSFDataValidation data_validation = new HSSFDataValidation(regions, constraint);
        return data_validation;
    }


    public static void addValuesToSheet(int startRow,int col,String[] values,HSSFSheet sheet) {
        for(int i = startRow; values != null && i < startRow + values.length; i++) {
            sheet.createRow(i).createCell(col).setCellValue(values[i-startRow]);
        }
    }

    private static String mapToEqualsString(Map<String,Object> map) {
        if(map == null || map.size()  <= 0)
            return "";
        StringBuilder builder = new StringBuilder();
        for(String key : map.keySet()) {
            builder.append(key+ " = "+map.get(key) +" \r\n");
        }
        return builder.toString();
    }
    /**
     * 将excel文件读为List<Map<列头名称,单元格值>>,支持xls,xlsx；空排行连续超过10行自动终止
     * @param fileName
     * @param startRow 列头开始的行号,行号从0开始
     * @return
     */
    public static List<Map<String,Object>> readFromExcelFile(String fileName,int startRow) {
        List<Map<String,Object>> datas = null;
        Workbook book = null;
        try {
            if(fileName.endsWith(".xls")) {
                book = new HSSFWorkbook(new FileInputStream(new File(fileName)));
            } else if(fileName.endsWith(".xlsx")) {
                book = new XSSFWorkbook(new FileInputStream(new File(fileName)));
            } else {
                return Collections.EMPTY_LIST;
            }
            Sheet sheet = book.getSheetAt(0);
            datas = iteratorSheetToList(sheet,startRow);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if(book != null) {
                    book.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return datas;
    }

    public static List<Map<String,Object>> readFromExcelFile(File file,int startRow) {
        List<Map<String,Object>> datas = null;
        Workbook book = null;
        try {
            if(file.getName().endsWith(".xls")) {
                book = new HSSFWorkbook(new FileInputStream(file));
            } else if(file.getName().endsWith(".xlsx")) {
                book = new XSSFWorkbook(new FileInputStream(file));
            } else {
                return Collections.EMPTY_LIST;
            }
            Sheet sheet = book.getSheetAt(0);
            datas = iteratorSheetToList(sheet,startRow);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if(book != null) {
                    book.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return datas;
    }

    public static List<Map<String,Object>> iteratorSheetToList(Sheet sheet,int startRow) {
        List<Map<String,Object>> datas = new ArrayList<Map<String,Object>>();
        Iterator<Row> it = sheet.rowIterator();
        Row firstRow = null;
        int continueBlankRow = 0;
        while (it.hasNext()) {
            Row row = (Row) it.next();
            if(row.getRowNum() < startRow ) {
                continue;
            }
            if(firstRow == null && row.getRowNum() == startRow) {
                firstRow = row;
                continue;
            }
            Map<String,Object> rowMap = new HashMap<String,Object>();
            int maxRow = firstRow!=null?firstRow.getPhysicalNumberOfCells():row.getPhysicalNumberOfCells();

            boolean continueBlankCell = true;

            for(int col = 0; col < maxRow; col++) {
                //跳过列头是空白的
                if(firstRow == null || firstRow.getCell(col) == null || TextUtils.isEmpty(firstRow.getCell(col).getStringCellValue())) continue;
                String columnName = firstRow!=null?(firstRow.getCell(col)!=null?firstRow.getCell(col).getStringCellValue():col+""):col+"";
                Object value = getCellValue(row.getCell(col));
                //该行全部列均为空值，才认为是空白行
                if(!TextUtils.isEmpty(value.toString())) {
                    continueBlankCell = false;
                }
                rowMap.put(columnName.trim(), value);
            }
            if(continueBlankCell) {
                continueBlankRow++;
            } else {
                datas.add(rowMap);
                continueBlankRow = 0;
            }
            //连续空白行大于10行认为结束
            if(continueBlankRow > 10) {
                break;
            }
        }
        return datas;
    }

    public static void writeToExcel(String fileName,String title,String[] columns,List<Map> datas) throws IOException {
        if(datas == null || datas.size() <= 0) return ;
        int rowSize = datas.size()+1;
        int colSize = columns.length;
        int rowNum = 2;
        File bookFile = new File(fileName);
        if(!bookFile.exists()) {
            bookFile.createNewFile();
        }
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = createSheetByTitles(workbook,title,columns);
        for(Map<String,Object> data : datas) {
            XSSFRow row = sheet.createRow(rowNum++);
            for(int col = 0; col < columns.length; col++ ) {
                row.createCell(col).setCellValue(data.get(columns[col]).toString());
            }
        }
        FileOutputStream fos = null;
        try {
            fos = new FileOutputStream(bookFile);
            workbook.write(fos);
        } catch (Exception e) {
        } finally {
            if(fos != null) {
                fos.flush();
                fos.close();
            }
        }
    }
}
