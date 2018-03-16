package com.yexj.excelOperator.dto;

import com.yexj.excelOperator.annotation.Column;
import com.yexj.excelOperator.annotation.Sheet;

import java.io.Serializable;
import java.util.Collection;

/**
 * Created by xinjian.ye on 2018/3/8.
 */
@Sheet
public class Employee implements Serializable {

    public Employee() {}

    public Employee(String name, Integer age, String email, String department) {
        this.email = email;
        this.age = age;
        this.name = name;
        this.department = department;
    }

    @Column(name = "姓名", priority = 0)
    private String name;

    @Column(name = "电子邮箱", priority = 2)
    private String email;

    @Column(name = "年龄", priority = 1)
    private Integer age;

    @Column(name = "部门", priority = 1)
    private String department;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getEmail() {
        return email;
    }

    public void setEmail(String email) {
        this.email = email;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    public String getDepartment() {
        return department;
    }

    public void setDepartment(String department) {
        this.department = department;
    }

    @Override
    public String toString() {
        return "Employee{" +
                "name='" + name + '\'' +
                ", email='" + email + '\'' +
                ", department =" + department +
                ", age=" + age +
                '}';
    }
}
