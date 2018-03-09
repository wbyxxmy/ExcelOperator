package com.yexj.excepOperator;

import com.yexj.excelOperator.Impl.FileServiceImpl;
import com.yexj.excelOperator.api.IFileService;
import com.yexj.excelOperator.dto.Employee;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Test;

import javax.annotation.Resource;
import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by xinjian.ye on 2018/3/8.
 */
public class FileServiceTest {
    IFileService fileService = new FileServiceImpl();

    @Test
    public void testExcelExprot() {

        List<Employee> employeeList = new ArrayList<Employee>();
        File file = new File("C:\\Users\\xinjian.ye\\Desktop\\personal\\ExcelOperator\\testExcelExport.xlsx");

        employeeList.add(new Employee("武则天1", 25, "wzt1@123.com"));
        employeeList.add(new Employee("武则天2", 26, "wzt2@123.com"));
        employeeList.add(new Employee("武则天3", 27, "wzt3@123.com"));
        employeeList.add(new Employee("武则天4", 28, "wzt4@123.com"));
        employeeList.add(new Employee("武则天5", 29, "wzt5@123.com"));
        employeeList.add(new Employee("武则天6", 30, "wzt6@123.com"));
        employeeList.add(new Employee("武则天7", 31, "wzt7@123.com"));

        employeeList.add(new Employee("罗永浩1", 25, "lyh1@123.com"));
        employeeList.add(new Employee("罗永浩2", 26, "lyh2@123.com"));
        employeeList.add(new Employee("罗永浩3", 27, "lyh3@123.com"));
        employeeList.add(new Employee("罗永浩4", 28, "lyh4@123.com"));
        employeeList.add(new Employee("罗永浩5", 29, "lyh5@123.com"));
        employeeList.add(new Employee("罗永浩6", 30, "lyh6@123.com"));
        employeeList.add(new Employee("罗永浩7", 31, "lyh7@123.com"));

        SXSSFWorkbook sxssfWorkbook = null;
        try {
            sxssfWorkbook = fileService.excelExport(employeeList);
            FileOutputStream out = new FileOutputStream(file);
            sxssfWorkbook.write(out);
        } catch (Exception e) {

        }
    }

    @Test
    public void testExcelRead() {
        try {
            fileService.excelRead();
        } catch (Exception e) {

        }
    }
}
