package com.excel.vo;


import com.jeff.regan.excel.annotation.ExcelField;

import java.util.Date;

/**
 * 学生 excel测试
 */
public class Student {
    private static final long serialVersionUID = -4026917215285783232L;

    @ExcelField(title = "姓名",sort = 1)
    private String name;
    @ExcelField(title = "学校" ,sort = 3)
    private String school;
    @ExcelField(title = "年龄",sort=2)
    private Integer age;
    @ExcelField(title = "入学时间",sort = 4)
    private Date joinDate;

    public Student() {
    }

    public Student(String name, String school, Integer age) {
        this.name = name;
        this.school = school;
        this.age = age;
        this.joinDate = new Date();
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getSchool() {
        return school;
    }

    public void setSchool(String school) {
        this.school = school;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    public Date getJoinDate() {
        return joinDate;
    }

    public void setJoinDate(Date joinDate) {
        this.joinDate = joinDate;
    }

    @Override
    public String toString() {
        return "Student{" +
                "name='" + name + '\'' +
                ", school='" + school + '\'' +
                ", age=" + age +
                ", joinDate=" + joinDate +
                '}';
    }
}
