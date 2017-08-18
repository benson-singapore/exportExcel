package com.jeff.regan.excel.vo;

import com.jeff.regan.excel.annotation.ExcelField;
import com.jeff.regan.excel.util.DateUtils;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;

/**
 * 工作簿
 *
 * @author zhangby
 * @date 2017/8/2 17:29
 */
public class ExcelSheet {
    public static final String GET_METHOD_TYPE = "get"; // 获取method
    private Sheet sheet;
    private ExcelTitle excelTitle; //excel 表格头部信息

    /**
     * 构造方法
     *
     * @param sheet
     */
    public ExcelSheet(Sheet sheet) {
        this.sheet = sheet;
    }

    /**
     * 获取row
     *
     * @param rownum
     * @return
     */
    public ExcelRow row(int rownum) {
        Row row = this.sheet.getRow(rownum);
        if (row == null) {
            row = this.sheet.createRow(rownum);
        }
        return new ExcelRow(row);
    }

    /**
     * 获取row
     *
     * @param rownum
     * @return
     */
    public ExcelRow getRow(int rownum) {
        return row(rownum);
    }

    public Sheet getHssfSheet() {
        return sheet;
    }

    /**
     * 合并单元格
     *
     * @param region
     * @return
     */
    public ExcelSheet addMergedRegion(CellRangeAddress region) {
        this.sheet.addMergedRegion(region);
        return this;
    }

    /**
     * 通过注解，导出excel
     *
     * @param clazz
     * @param cellStyle
     * @return
     */
    public ExcelHeader header(Class<?> clazz, CellStyle cellStyle) throws NoSuchMethodException, IllegalAccessException, InvocationTargetException {
        Field[] fields = getAllFields(clazz); //获取所有属性
        List<Field> rsFields = sortFields(fields);
        int row = 0;
        if (this.excelTitle != null && this.excelTitle.title != null) {
            row = 1;
        }
        //设置头文件
        for (int i = 0; i < rsFields.size(); i++) {
            Field field = rsFields.get(i);
            ExcelField annotation = field.getAnnotation(ExcelField.class);
            String title = annotation.title();
            ExcelRow excelRow = this.row(row).cell(i).cellValue(title);
            if (cellStyle != null) {
                excelRow.cellStyle(cellStyle);
            }
        }
        //调用生成list数据
        return new ExcelHeader(this, row);
    }

    /**
     * 通过注解，导出excel
     *
     * @param clazz
     * @return
     */
    public ExcelHeader header(Class<?> clazz) throws NoSuchMethodException, IllegalAccessException, InvocationTargetException {
        return header(clazz, null);
    }


    /**
     * 设置title
     *
     * @param excelTitle
     * @return
     */
    public ExcelTitle title(String excelTitle) {
        ExcelTitle excelTitle1 = new ExcelTitle(excelTitle);
        this.excelTitle = excelTitle1;
        return excelTitle1;
    }

    /**
     * 设置excel 循环list
     *
     * @param list  list数据
     * @param rowStart row开始行
     * @param cellStart cell 开始行
     * @return
     */
    public <E> CellData setDateList(List<E> list, Integer rowStart, Integer cellStart) throws NoSuchMethodException, IllegalAccessException, InvocationTargetException {
        return setDateList(list, rowStart, null, cellStart, null);
    }

    public <E> CellData setDateList(List<E> list, Integer rowStart, Integer rowEnd, Integer cellStart, Integer cellEnd) throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {
        List<ExcelRow> rowList = new ArrayList<>();
        if (list.size() < 1) return new CellData(rowList);
        Class<?> clazz = list.get(0).getClass();
        Field[] fields = getAllFields(clazz); //获取所有属性
        //对参与excel导出的序列，进行排序
        List<Field> rsFields = sortFields(fields);

        //设置excel 值
        for (E e : list) {
            int cellNum = cellStart;
            for (Field fd : rsFields) {
                if ((rowEnd != null && rowStart > rowEnd) || (cellEnd != null && cellEnd > cellNum)) {
                    break;
                }
                //获取值
                Method method = clazz.getMethod(toGetMethod(fd.getName(), GET_METHOD_TYPE));
                ExcelRow cell = this.row(rowStart).cell(cellNum);
                rowList.add(cell);
                addCell(cell, fd.getType(), method.invoke(e));
                cellNum++;
            }
            rowStart += 1;
        }
        autoWeight(rsFields.size());
        return new CellData(rowList);
    }

    /**
     * 对注解字段，按照sort排序
     * @param fields
     * @return
     */
    public List<Field> sortFields(Field[] fields) {
        //对参与excel导出的序列，进行排序
        List<Field> rsFields = Arrays.asList(fields).stream().filter(v -> {
            //过滤掉不包含注解的属性
            return (v.getAnnotation(ExcelField.class) != null) ? true : false;
        }).sorted((v1, v2) -> {
            //对属性按照排序方式，重新排列
            int sort1 = v1.getAnnotation(ExcelField.class).sort();
            int sort2 = v2.getAnnotation(ExcelField.class).sort();
            return String.valueOf(sort1).compareTo(String.valueOf(sort2));
        }).collect(Collectors.toList());
        return rsFields;
    }

    /**
     * 自动设置宽度
     *
     * @param columnNum
     */
    public void autoWeight(int columnNum) {
        //让列宽随着导出的列长自动适应
        for (int colNum = 0; colNum < columnNum; colNum++) {
            int columnWidth = sheet.getColumnWidth(colNum) / 256;
            for (int rowNum = 0; rowNum < sheet.getLastRowNum(); rowNum++) {
                Row currentRow;
                //当前行未被使用过
                if (sheet.getRow(rowNum) == null) {
                    currentRow = sheet.createRow(rowNum);
                } else {
                    currentRow = sheet.getRow(rowNum);
                }
                if (currentRow.getCell(colNum) != null) {
                    Cell currentCell = currentRow.getCell(colNum);
                    if (currentCell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
                        int length = currentCell.getStringCellValue().getBytes().length;
                        if (columnWidth < length) {
                            columnWidth = length;
                        }
                    }
                }
            }
            sheet.setColumnWidth(colNum, (columnWidth + 4) * 256);
        }

        /**
         * 合并标题
         */
        if (this.excelTitle != null && this.excelTitle.title != null) {
            this.sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, columnNum - 1));
            ExcelRow excelRow = this.row(0).cell(0).cellValue(this.excelTitle.title);
            excelRow.getRow().setHeightInPoints(30.0F);
            if (this.excelTitle.cellStyle != null) {
                excelRow.cellStyle(this.excelTitle.cellStyle);
            }
        }
    }

    /**
     * 反射获取类中的全部属性，包括父类中的属性。
     * @param clazz
     * @return
     */
    public static Field[] getAllFields(Class<?> clazz) {
        Field[] fileds = new Field[0];
        fileds = (Field[]) ArrayUtils.addAll(clazz.getDeclaredFields(), new Field[0]);
        Class<?> parent = clazz.getSuperclass();
        if (null != parent) {
            fileds = (Field[]) ArrayUtils.addAll(fileds, getAllFields(parent));
        }

        return fileds;
    }

    /**
     * 根据反射类型设置 cell
     * @param excelRow row对象
     * @param fieldType 反射对象
     * @param val 值
     */
    private <E> void addCell(ExcelRow excelRow, Class<?> fieldType, Object val) {
        try {
            if (val == null) {
                excelRow.cellValue("");
            } else if (val instanceof String) {
                excelRow.cellValue((String) val);
            } else if (val instanceof Integer) {
                excelRow.cellValue((double) ((Integer) val).intValue());
            } else if (val instanceof Long) {
                excelRow.cellValue((double) ((Long) val).longValue());
            } else if (val instanceof Double) {
                excelRow.cellValue(((Double) val).doubleValue());
            } else if (val instanceof Float) {
                excelRow.cellValue((double) ((Float) val).floatValue());
            } else if (val instanceof Date) {
                excelRow.cellValue(new SimpleDateFormat("yyyy-MM-dd").format(val));
            } else if (fieldType != Class.class) {
                excelRow.cellValue((String) fieldType.getMethod("setValue", Object.class).invoke((Object) null, val));
            } else {
                excelRow.cellValue((String) Class.forName(this.getClass().getName().replaceAll(this.getClass().getSimpleName(), "fieldtype." + val.getClass().getSimpleName() + "Type")).getMethod("setValue", Object.class).invoke((Object) null, val));
            }
        } catch (Exception var9) {
            if(val != null){
                excelRow.cellValue(val.toString());
            }
        }
    }

    /**
     * excel 导入
     *
     * @param rowNum  row 开始位置
     * @param cellNum cell 开始位置
     * @param keys    格式："姓名:name,年龄:age"
     * @return
     */
    public ImportExcel getList(int rowNum, int cellNum, String... keys) {
        List<Map<String, String>> rsList = new ArrayList<>();
        //获取header
        List<String> headers = getExcelHeader(rowNum, cellNum, keys);
        //循环解析data数据
        rowNum++;
        for (int i = rowNum; this.sheet.getRow(i) != null; i++) {
            Map<String, String> data = new LinkedHashMap<>(); //建立有序的Map
            for (int j = cellNum; this.sheet.getRow(i).getCell(j) != null; j++) {
                Cell cell = this.sheet.getRow(i).getCell(j);
                cell.setCellType(Cell.CELL_TYPE_STRING);
                data.put(headers.get(j), cell.getStringCellValue());
            }
            rsList.add(data);
        }
        return new ImportExcel(rsList, rowNum, cellNum, keys);
    }

    /**
     * 获取header
     * @param rowNum row 起始位置
     * @param cellNum cell起始位置
     * @param keys 自定义导出数据方式
     * @return
     */
    private List<String> getExcelHeader(int rowNum, int cellNum, String[] keys) {
        List<String> headers = new ArrayList<>();
        List<String> rsHeaderList = new ArrayList<>();
        for (int i = cellNum; this.sheet.getRow(rowNum).getCell(i) != null; i++) {
            Cell cell = this.sheet.getRow(rowNum).getCell(i);
            cell.setCellType(Cell.CELL_TYPE_STRING);
            headers.add(cell.getStringCellValue());
        }
        if (keys != null && keys.length > 0) {
            //原数据替换
            try {
                String[] keyArr = keys[0].split(",");
                Map<String, String> map = new HashMap<>();
                for (String key : keyArr) {
                    String[] split = key.split(":");
                    map.put(split[0], split[1]);
                }
                for (int i = 0; i < headers.size(); i++) {
                    String hd = map.get(headers.get(i));
                    if (hd != null) {
                        rsHeaderList.add(hd);
                    } else {
                        rsHeaderList.add(headers.get(i));
                    }
                }
            } catch (Exception e) {
                throw new RuntimeException("keys 格式异常!");
            }
        } else {
            rsHeaderList = headers;
        }
        return rsHeaderList;
    }

    /**
     * 把属性转换为get方法
     *
     * @param filed
     * @return "bmys" -> "getBmys"
     */
    public static String toGetMethod(String filed, String methodType) {
        return methodType + filed.substring(0, 1).toUpperCase() + filed.substring(1, filed.length());
    }

    /**
     * excel row
     *
     * @author zhangby
     * @date 2017/8/2 17:29
     */
    public class ExcelRow {
        private Row row;
        private Cell cell;

        public ExcelRow(Row row) {
            this.row = row;
        }

        public ExcelRow cell(int cellnum) {
            this.cell = this.row.createCell(cellnum);
            return this;
        }

        //设置样式
        public ExcelRow cellStyle(HSSFCellStyle hssfCellStyle) {
            this.cell.setCellStyle(hssfCellStyle);
            return this;
        }

        public ExcelRow cellStyle(XSSFCellStyle hssfCellStyle) {
            this.cell.setCellStyle(hssfCellStyle);
            return this;
        }

        public ExcelRow cellStyle(CellStyle cellStyle) {
            this.cell.setCellStyle(cellStyle);
            return this;
        }

        public ExcelRow cellValue(String cellValue) {
            this.cell.setCellValue(cellValue);
            return this;
        }

        public ExcelRow cellValue(boolean cellValue) {
            this.cell.setCellValue(cellValue);
            return this;
        }

        public ExcelRow cellValue(Calendar cellValue) {
            this.cell.setCellValue(cellValue);
            return this;
        }

        public ExcelRow cellValue(RichTextString cellValue) {
            this.cell.setCellValue(cellValue);
            return this;
        }

        public ExcelRow cellValue(Date cellValue) {
            this.cell.setCellValue(cellValue);
            return this;
        }

        public ExcelRow cellValue(double cellValue) {
            this.cell.setCellValue(cellValue);
            return this;
        }

        public Row getRow() {
            return row;
        }

        private Cell getCell() {
            return cell;
        }

        public ExcelRow getCell(int cellNum) {
            ExcelRow excelRow = new ExcelRow(this.row);
            excelRow.cell = excelRow.getRow().getCell(cellNum);
            return excelRow;
        }

        public String getCellValue(){
            if(this.cell == null) return null;
            this.cell.setCellType(Cell.CELL_TYPE_STRING);
            return this.cell.getStringCellValue();
        }
    }

    /**
     * 统一操作excel cell数据
     */
    public class CellData {
        List<ExcelRow> rowList = new ArrayList<>(); //对属性操作

        public CellData(List<ExcelRow> rowList) {
            this.rowList = rowList;
        }

        /**
         * 统一设置样式
         *
         * @param cellStyle
         * @return
         */
        public CellData cellStyle(CellStyle cellStyle) {
            if (rowList.size() > 0) {
                rowList.forEach(row -> {
                    row.getCell().setCellStyle(cellStyle);
                });
            }
            return this;
        }

        /**
         * 统一设置样式
         *
         * @param cellStyle
         * @return
         */
        public CellData cellStyle(HSSFCellStyle cellStyle) {
            if (rowList.size() > 0) {
                rowList.forEach(row -> {
                    row.getCell().setCellStyle(cellStyle);
                });
            }
            return this;
        }

        /**
         * 统一设置样式
         *
         * @param cellStyle
         * @return
         */
        public CellData cellStyle(XSSFCellStyle cellStyle) {
            if (rowList.size() > 0) {
                rowList.forEach(row -> {
                    row.getCell().setCellStyle(cellStyle);
                });
            }
            return this;
        }

        public List<ExcelRow> getList() {
            return rowList;
        }
    }

    /**
     *
     */
    public class ExcelHeader {
        private ExcelSheet excelSheet;
        private int row;

        public ExcelHeader(ExcelSheet excelSheet, int row) {
            this.excelSheet = excelSheet;
            this.row = row;
        }

        public <E> CellData setData(List<E> list) throws NoSuchMethodException, IllegalAccessException, InvocationTargetException {
            return excelSheet.setDateList(list, row + 1, 0);
        }
    }

    /**
     * 标题
     */
    public class ExcelTitle {
        private String title;
        private CellStyle cellStyle;

        public ExcelTitle(String exctelTitle) {
            this.title = exctelTitle;
        }

        public ExcelTitle cellStyle(CellStyle cellStyle) {
            this.cellStyle = cellStyle;
            return this;
        }

        public ExcelTitle cellStyle(HSSFCellStyle cellStyle) {
            this.cellStyle = cellStyle;
            return this;
        }

        public ExcelTitle cellStyle(XSSFCellStyle cellStyle) {
            this.cellStyle = cellStyle;
            return this;
        }
    }

    /**
     * excel 导入工具
     */
    public class ImportExcel {
        List<Map<String, String>> rsList = new ArrayList<>();
        int rowNum; //row起始位置
        int cellNum; //cell起始位置
        String[] keys;

        /**
         * 构造方法
         *
         * @param rsList
         */
        public ImportExcel(List<Map<String, String>> rsList, int rowNum, int cellNum, String... keys) {
            this.rsList = rsList;
            this.rowNum = rowNum;
            this.cellNum = cellNum;
            this.keys = keys;
        }

        /**
         * 数据转Map
         *
         * @return
         */
        public List<Map<String, String>> toMap() {
            return this.rsList;
        }

        /**
         * 依据注解转Map
         * @return
         */
        public <T> List<Map<String, String>> toMap4Annotation(Class<T> clazz) {
            //获取 excel header
            List<String> headers = new ArrayList<>();
            for (int i = cellNum; sheet.getRow(rowNum).getCell(i) != null; i++) {
                Cell cell = sheet.getRow(rowNum).getCell(i);
                cell.setCellType(Cell.CELL_TYPE_STRING);
                headers.add(cell.getStringCellValue());
            }
            //获取对象注解title与属性name值
            List<Field> rsFields = sortFields(getAllFields(clazz));
            Map<String, String> headerMap = new HashMap<>();
            for (int i=0;i<rsFields.size();i++) {
                Field field = rsFields.get(i);
                ExcelField annotation = field.getAnnotation(ExcelField.class);
                if (annotation != null) {
                    headerMap.put(annotation.title(), field.getName());
                }else{
                    String hkey = headers.get(i)!=null?headers.get(i):field.getName();
                    headerMap.put(hkey,field.getName());
                }
            }

            List<Map<String, String>> dataList = new ArrayList<>();
            //data替换key值替换
            rsList.forEach(map->{
                Map<String,String> rsMap = new LinkedHashMap<>();
                map.forEach((key,val)->{
                    rsMap.put(headerMap.get(key), val);
                });
                dataList.add(rsMap);
            });

            return dataList;
        }


        /**
         * 转 class
         *
         * @param <T>
         * @param clazz
         * @return
         */
        public <T> List<T> toObject(Class<T> clazz) {
            List<T> list = new ArrayList<>();
            if (keys != null && keys.length > 0) {
                this.rsList.forEach(m -> {
                    try {
                        list.add(mapToEntity(m, clazz));
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                });
            } else {
                //转对象
                toMap4Annotation(clazz).forEach(m -> {
                    try {
                        list.add(mapToEntity(m, clazz));
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                });
            }
            return list;
        }
    }

    /**
     * map 转 entity方法
     * @return
     */
    public  <T> T mapToEntity(Map<String, String> map,Class<?> clazz) throws IllegalAccessException, InstantiationException, ParseException {
        T entity = (T) clazz.newInstance();
        //对参与excel导出的序列，进行排序
        List<Field> fields = sortFields(getAllFields(clazz));
        // 循环查询出的列
        for (Field field : fields) {
            String fName = field.getName();
            if (!map.containsKey(fName)) {
                fName = fName.toUpperCase();// JDBC方式查询出的别名都是大写
            }
            if (map.containsKey(fName)) {
                field.setAccessible(true);
                Object obj = null;
                if(field.getType() == String.class){
                    obj = map.get(fName);
                }else if(field.getType() == Integer.class){
                    obj = Integer.parseInt(map.get(fName));
                }else if(field.getType() == Long.class){
                    obj = Long.parseLong(map.get(fName));
                }else if(field.getType() == Double.class){
                    obj = Double.parseDouble(map.get(fName));
                }else if(field.getType() == Float.class){
                    obj = Float.parseFloat(map.get(fName));
                }else if(field.getType() == Date.class){
                    obj = DateUtils.str2Date(map.get(fName));
                }
                field.set(entity, obj);
            }
        }
        return entity;
    }

}