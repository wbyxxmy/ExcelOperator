package com.yexj.excepOperator;

import com.yexj.excelOperator.Impl.FileServiceImpl;
import com.yexj.excelOperator.api.IFileService;
import com.yexj.excelOperator.dto.Employee;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Test;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
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

        employeeList.add(new Employee("武则天1", 25, "wzt1@123.com", "政治部"));
        employeeList.add(new Employee("武则天2", 26, "wzt2@123.com", "政治部"));
        employeeList.add(new Employee("武则天3", 27, "wzt3@123.com", "政治部"));
        employeeList.add(new Employee("武则天4", 28, "wzt4@123.com", "政治部"));
        employeeList.add(new Employee("武则天5", 29, "wzt5@123.com", "政治部"));
        employeeList.add(new Employee("武则天6", 30, "wzt6@123.com", "政治部"));
        employeeList.add(new Employee("武则天7", 31, "wzt7@123.com", "政治部"));

        employeeList.add(new Employee("罗永浩1", 25, "lyh1@123.com", "营运部"));
        employeeList.add(new Employee("罗永浩2", 26, "lyh2@123.com", "营运部"));
        employeeList.add(new Employee("罗永浩3", 27, "lyh3@123.com", "营运部"));
        employeeList.add(new Employee("罗永浩4", 28, "lyh4@123.com", "营运部"));
        employeeList.add(new Employee("罗永浩5", 29, "lyh5@123.com", "营运部"));
        employeeList.add(new Employee("罗永浩6", 30, "lyh6@123.com", "营运部"));
        employeeList.add(new Employee("罗永浩7", 31, "lyh7@123.com", "营运部"));

        SXSSFWorkbook sxssfWorkbook = null;
        try {
            sxssfWorkbook = fileService.excelExport(employeeList);
            FileOutputStream out = new FileOutputStream(file);
            sxssfWorkbook.write(out);
        } catch (Exception e) {

        }
    }

    @Test
    public void testExcelUplode() throws Exception {
        File file = new File("C:\\Users\\xinjian.ye\\Desktop\\personal\\ExcelOperator\\testExcelExport.xlsx");
        MultipartFile multipart = null;
        try {
            FileInputStream inputStream = new FileInputStream(file);
            multipart = new MockMultipartFile(file.getName(),"testExcelExport.xlsx","xlsx",inputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }

//        LeadingInExcel<Employee> testExcel=null;
        List<?> uploadAndRead=null;
        boolean judgement = false;
        String Msg =null;
        String error = "";


        //定义需要读取的数据
        String formart = "yyyy-MM-dd";
        String propertiesFileName = "config";
        String kyeName = "file_path";
        int sheetIndex = 0;


        //调用解析工具包
//        testExcel=new LeadingInExcel<Employee>(formart);
        //解析excel，获取客户信息集合
        uploadAndRead = fileService.uploadAndRead(multipart, propertiesFileName, kyeName, sheetIndex, Employee.class);
        if(uploadAndRead != null && !"[]".equals(uploadAndRead.toString()) && uploadAndRead.size()>=1){
            judgement = true;
        }
        if(judgement){

            //把客户信息分为没100条数据为一组迭代添加客户信息（注：将customerList集合作为参数，在Mybatis的相应映射文件中使用foreach标签进行批量添加。）
            //int count=0;
            int listSize=uploadAndRead.size();
            int toIndex=100;
            for (int i = 0; i < listSize; i+=100) {
                if(i+100>listSize){
                    toIndex=listSize-i;
                }
                List<?> subList = uploadAndRead.subList(i, i+toIndex);
                /** 此处执行集合添加 */
                for(Object o : subList) {
                    System.out.println(o);
                }
            }

            Msg ="批量导入EXCEL成功！";
        }else{
            Msg ="批量导入EXCEL失败！";
        }

        String res = "{ error:'" + error + "', msg:'" + Msg + "'}";
//        return res;
    }

    @Test
    public void testReflect() throws Exception {
        Class clazz = Employee.class;
        Object employ = clazz.newInstance();

        System.out.println(employ);
    }
}
