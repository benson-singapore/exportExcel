package com.excel;

import com.excel.vo.Student;
import com.jeff.regan.excel.util.ExcelUtil;
import com.jeff.regan.excel.vo.Excel;
import com.jeff.regan.excel.vo.ExcelSheet;
import org.junit.Test;

import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.List;

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
     * excel数据导入 （无需注解）
     */
    @Test
    public void imortExcelNoAn() throws Exception {
        String keyValue = "姓名:name,学校:school,年龄:age,入学时间:joinDate";
        List<Object> students = ExcelUtil.readXlsPart("c://student_bak.xlsx", ExcelUtil.getMap(keyValue), Student.class,2);
        students.forEach(s->System.out.println(s));
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
}
