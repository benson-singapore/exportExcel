package com.jeff.regan.excel.vo;

import com.jeff.regan.excel.annotation.ExcelField;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.util.CellRangeAddress;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
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
    private HSSFSheet hssfSheet;

    public ExcelSheet(HSSFSheet hssfSheet) {
        this.hssfSheet = hssfSheet;
    }

    public ExcelRow row(int rownum) {
        HSSFRow row = this.hssfSheet.getRow(rownum);
        if (row == null) {
            row = this.hssfSheet.createRow(rownum);
        }
        return new ExcelRow(row);
    }

    public HSSFSheet getHssfSheet() {
        return hssfSheet;
    }

    public ExcelSheet addMergedRegion(Region region){
        this.hssfSheet.addMergedRegion(region);
        return this;
    }
    public ExcelSheet addMergedRegion(CellRangeAddress region){
        this.hssfSheet.addMergedRegion(region);
        return this;
    }

    public <E> CellData setDateList(List<E> list, Integer rowStart, Integer cellStart) throws NoSuchMethodException, IllegalAccessException, InvocationTargetException {
        return setDateList(list, rowStart, null, cellStart, null);
    }

    public <E> CellData setDateList(List<E> list, Integer rowStart, Integer rowEnd, Integer cellStart, Integer cellEnd) throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {
        List<ExcelRow> rowList = new ArrayList<>();
        if (list.size() < 1) new CellData(rowList);
        Class<?> clazz = list.get(0).getClass();
        Field[] fields = getAllFields(clazz); //获取所有属性
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

    public void autoWeight(int columnNum) {
        //让列宽随着导出的列长自动适应
        for (int colNum = 0; colNum < columnNum; colNum++) {
            int columnWidth = hssfSheet.getColumnWidth(colNum) / 256;
            for (int rowNum = 0; rowNum < hssfSheet.getLastRowNum(); rowNum++) {
                HSSFRow currentRow;
                //当前行未被使用过
                if (hssfSheet.getRow(rowNum) == null) {
                    currentRow = hssfSheet.createRow(rowNum);
                } else {
                    currentRow = hssfSheet.getRow(rowNum);
                }
                if (currentRow.getCell(colNum) != null) {
                    HSSFCell currentCell = currentRow.getCell(colNum);
                    if (currentCell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
                        int length = currentCell.getStringCellValue().getBytes().length;
                        if (columnWidth < length) {
                            columnWidth = length;
                        }
                    }
                }
            }
            hssfSheet.setColumnWidth(colNum, (columnWidth + 4) * 256);
        }
    }

    public static Field[] getAllFields(Class<?> clazz) {
        Field[] fileds = new Field[0];
        fileds = (Field[]) ArrayUtils.addAll(clazz.getDeclaredFields(), new Field[0]);
        Class<?> parent = clazz.getSuperclass();
        if (null != parent) {
            fileds = (Field[])ArrayUtils.addAll(fileds, getAllFields(parent));
        }

        return fileds;
    }

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
            excelRow.cellValue(val.toString());
        }
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
        private HSSFRow row;
        private HSSFCell cell;

        public ExcelRow(HSSFRow row) {
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

        public HSSFRow getRow() {
            return row;
        }

        public HSSFCell getCell() {
            return cell;
        }
    }

    /**
     * 统一操作excel cell数据
     */
    public class CellData{
        List<ExcelRow> rowList = new ArrayList<>(); //对属性操作

        public CellData(List<ExcelRow> rowList) {
            this.rowList = rowList;
        }

        /**
         * 统一设置样式
         * @param cellStyle
         * @return
         */
        public CellData cellStyle(CellStyle cellStyle){
            if (rowList.size() > 0) {
                rowList.forEach(row->{
                    row.getCell().setCellStyle(cellStyle);
                });
            }
            return this;
        }
        /**
         * 统一设置样式
         * @param cellStyle
         * @return
         */
        public CellData cellStyle(HSSFCellStyle cellStyle){
            if (rowList.size() > 0) {
                rowList.forEach(row->{
                    row.getCell().setCellStyle(cellStyle);
                });
            }
            return this;
        }

        public List<ExcelRow> getList() {
            return rowList;
        }
    }
}