package com.excel;

import com.excel.vo.Student;
import com.jeff.regan.excel.vo.Excel;
import com.jeff.regan.excel.vo.ExcelSheet;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * excel 导入导出测试类
 *
 * @author zhangby
 * @date 2017/8/2 15:09
 */
public class ExcelTest {

    /**
     * excel 新建
     * @throws IOException
     */
    @Test
    public void excelTest() throws IOException {
        Excel excel = new Excel();
        ExcelSheet sheet = excel.createSheet();
        sheet.row(0).cell(2).cellValue("1");
        sheet.row(1).cell(2).cellValue("2");
        excel.saveExcel("c://test1.xlsx");
    }

    /**
     * 模板部分数据替换
     * @throws IOException
     */
    @Test
    public void excelModel() throws IOException {
        Excel excel = new Excel("c://test1.xlsx");
        ExcelSheet sheet = excel.getSheet(); //默认获取第一个工作簿
        sheet.row(0).cell(2).cellValue("111111111");
        excel.saveExcel("c://test2.xlsx");
    }

    /**
     * 模板指定位置 list数据循环导出（需要entity注解）
     * @throws IOException
     * @throws NoSuchMethodException
     * @throws IllegalAccessException
     * @throws InvocationTargetException
     */
    @Test
    public void myExcel() throws IOException, NoSuchMethodException, IllegalAccessException, InvocationTargetException {
        Excel excel = new Excel("c://student_temp.xls");
        ExcelSheet sheet = excel.getSheet();
        sheet.setDateList(init(), 2, 0);
        excel.saveExcel("c://student_temp_rs.xlsx");
    }

    /**
     * 通过注解导出Excel
     */
    @Test
    public void Export4Annotation() throws NoSuchMethodException, IllegalAccessException, InvocationTargetException, IOException {
        Excel excel = new Excel();
        ExcelSheet sheet = excel.createSheet();
        sheet.title("学生统计表"); //设置excel title名称
        sheet.header(Student.class).setData(init()); //设置 data
        excel.saveExcel("c://student_annotation.xlsx");
    }

    /**
     * 通过注解导出Excel
     */
    @Test
    public void Export4Annotation2Style() throws NoSuchMethodException, IllegalAccessException, InvocationTargetException, IOException {
        Excel excel = new Excel();
        ExcelSheet sheet = excel.createSheet();
        //获取excel样式
        Map<String, CellStyle> styles = createStyles(excel.getWorkbook());
        sheet.title("学生统计表").cellStyle(styles.get("title"));   //设置excel title名称
        sheet.header(Student.class, styles.get("header"))                    //设置excel hear
                .setData(init()).cellStyle(styles.get("data"));              //设置 data
        excel.saveExcel("c://student_annotation.xlsx");
    }

    @Test
    public void importExcelForMap1() throws IOException {
        Excel excel = new Excel("c://student_annotation.xlsx");
        ExcelSheet sheet = excel.getSheet();
        List<Map<String, String>> list = sheet.getList(1, 0).toMap();
        list.forEach(map -> System.out.println(map));
    }

    @Test
    public void importExcelForMap2() throws IOException {
        Excel excel = new Excel("c://student_annotation.xlsx");
        ExcelSheet sheet = excel.getSheet();
        List<Map<String, String>> list = sheet.getList(1, 0).toMap4Annotation(Student.class);
        list.forEach(map -> System.out.println(map));
    }

    @Test
    public void importExcelForMap3() throws IOException {
        String keyValue = "姓名:name,学校:school,年龄:age,入学时间:joinDate";
        Excel excel = new Excel("c://student_annotation.xlsx");
        ExcelSheet sheet = excel.getSheet();
        List<Map<String, String>> list = sheet.getList(1, 0,keyValue).toMap();
        list.forEach(map -> System.out.println(map));
    }

    @Test
    public void importExcelForObject() throws IOException {
        String keyValue = "姓名:name,学校:school,年龄:age,入学时间:joinDate";
        Excel excel = new Excel("c://student_annotation.xlsx");
        ExcelSheet sheet = excel.getSheet();
        List<Student> list = sheet.getList(1, 0,keyValue).toObject(Student.class);
        list.forEach(map -> System.out.println(map));
    }

    @Test
    public void importExcelForObject2Annotation() throws IOException {
        Excel excel = new Excel("c://student_annotation.xlsx");
        ExcelSheet sheet = excel.getSheet();
        List<Student> list = sheet.getList(1, 0).toObject(Student.class);
        list.forEach(map -> System.out.println(map));
    }

    /**
     * 初始化数据
     * @return
     */
    static List<Student>  init(){
        List<Student> list = new ArrayList<>();

        Student st1 = new Student("tom","huax",10);
        Student st2 = new Student("tom","huax",10);
        Student st3 = new Student("tom","huax",10);
        list.add(st1);
        list.add(st2);
        list.add(st3);
        list.forEach(s->System.out.println(s));
        return list;
    }

    /**
     * excel 样式
     * @return
     */
    public Map<String, CellStyle> createStyles(Workbook workbook) {
        Map<String, CellStyle> styles = new HashMap();
        CellStyle style = workbook.createCellStyle();
        style.setAlignment((short)2);
        style.setVerticalAlignment((short)1);
        Font titleFont = workbook.createFont();
        titleFont.setFontName("Arial");
        titleFont.setFontHeightInPoints((short)16);
        titleFont.setBoldweight((short)700);
        style.setFont(titleFont);
        styles.put("title", style);
        style = workbook.createCellStyle();
        style.setVerticalAlignment((short)1);
        style.setBorderRight((short)1);
        style.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderLeft((short)1);
        style.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderTop((short)1);
        style.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderBottom((short)1);
        style.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        Font dataFont = workbook.createFont();
        dataFont.setFontName("Arial");
        dataFont.setFontHeightInPoints((short)10);
        style.setFont(dataFont);
        styles.put("data", style);
        style = workbook.createCellStyle();
        style.cloneStyleFrom((CellStyle)styles.get("data"));
        style.setAlignment((short)1);
        styles.put("data1", style);
        style = workbook.createCellStyle();
        style.cloneStyleFrom((CellStyle)styles.get("data"));
        style.setAlignment((short)2);
        styles.put("data2", style);
        style = workbook.createCellStyle();
        style.cloneStyleFrom((CellStyle)styles.get("data"));
        style.setAlignment((short)3);
        styles.put("data3", style);
        style = workbook.createCellStyle();
        style.cloneStyleFrom((CellStyle)styles.get("data"));
        style.setAlignment((short)2);
        style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setFillPattern((short)1);
        Font headerFont = workbook.createFont();
        headerFont.setFontName("Arial");
        headerFont.setFontHeightInPoints((short)10);
        headerFont.setBoldweight((short)700);
        headerFont.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(headerFont);
        styles.put("header", style);
        return styles;
    }
}
