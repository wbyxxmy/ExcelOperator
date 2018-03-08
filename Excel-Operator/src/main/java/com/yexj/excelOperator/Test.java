package com.yexj.excelOperator;

import com.yexj.excelOperator.Impl.FileServiceImpl;
import com.yexj.excelOperator.annotation.Column;
import com.yexj.excelOperator.annotation.Sheet;
import com.yexj.excelOperator.api.IFileService;

/**
 * Created by xinjian.ye on 2018/3/8.
 */
public class Test {
    public static void main(String[] args) throws Exception {
        IFileService fileService = new FileServiceImpl();

//        List<Employee> employeeList = new ArrayList<Employee>();
//        File file = new File("C:\\Users\\xinjian.ye\\Desktop\\personal\\ExcelOperator\\testExcelExport.xlsx");
//
//        employeeList.add(new Employee("武则天1", 25, "wzt1@123.com"));
//        employeeList.add(new Employee("武则天2", 26, "wzt2@123.com"));
//        employeeList.add(new Employee("武则天3", 27, "wzt3@123.com"));
//        employeeList.add(new Employee("武则天4", 28, "wzt4@123.com"));
//        employeeList.add(new Employee("武则天5", 29, "wzt5@123.com"));
//        employeeList.add(new Employee("武则天6", 30, "wzt6@123.com"));
//        employeeList.add(new Employee("武则天7", 31, "wzt7@123.com"));
//
//        SXSSFWorkbook sxssfWorkbook = fileService.excelExport(employeeList);
//        FileOutputStream out = new FileOutputStream(file);
//        sxssfWorkbook.write(out);

        fileService.excelRead();
    }

    @Sheet
    static class Employee {

        Employee(String name, Integer age, String email) {
            this.email = email;
            this.age = age;
            this.name = name;
        }
        @Column(name = "姓名", priority = 0)
        private String name;

        @Column(name = "电子邮箱", priority = 2)
        private String email;

        @Column(name = "年龄", priority = 1)
        private Integer age;

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
    }
}
