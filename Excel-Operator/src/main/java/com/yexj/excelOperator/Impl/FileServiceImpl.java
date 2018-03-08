package com.yexj.excelOperator.Impl;

import com.yexj.excelOperator.api.IFileService;
import com.yexj.excelOperator.utils.WorkbookUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.util.List;

/**
 * Created by xinjian.ye on 2018/3/7.
 */
public class FileServiceImpl implements IFileService {

    private static final Logger logger = LoggerFactory.getLogger(FileServiceImpl.class);

    /**
     * POI方式导出为Excel
     * @param dataInfoList
     * @return
     * @throws Exception
     */
    @Override
    public SXSSFWorkbook excelExport(List<?> dataInfoList) throws Exception {
        Class clazz = dataInfoList.get(0).getClass();
        SXSSFWorkbook workbook = WorkbookUtils.createWorkbook(clazz, dataInfoList);
        return workbook;
//        long a = System.currentTimeMillis();
//        SXSSFWorkbook workbook = new SXSSFWorkbook(10000);
//        logger.info("\r<br> new SXSSFWorkbook执行耗时 : "+(System.currentTimeMillis()-a)/1000f+" 秒 ");
//        workbook.setCompressTempFiles(true);
//        Sheet first = workbook.createSheet("Sheet1");
//        Class classes = dataInfoList.get(0).getClass();
//        int resultSize = dataInfoList.size();
//        for (int i = -1; i < resultSize; i++) {
//            if(i == -1) {
//                // 首行
//                Row row = first.createRow(0);
//                CellStyle style = workbook.createCellStyle();
//                Method method = classes.getMethod("setExcelHead", CellStyle.class, Row.class);
//                method.invoke(dataInfoList.get(0), style, row);
//            } else {
//                // 数据
//                Row row = first.createRow(i+1);
//                Method method = classes.getMethod("setExcelData", Row.class);
//                method.invoke(dataInfoList.get(i), row);
//            }
//        }
//        logger.info("\r<br> generate workbook执行耗时 : "+(System.currentTimeMillis()-a)/1000f+" 秒 ");
//        return workbook;
    }

    /**
     *
     * @throws Exception
     */
    @Override
    public void excelRead() throws Exception {
        XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(new File("C:\\Users\\xinjian.ye\\Desktop\\personal\\ExcelOperator\\testExcelExport.xlsx")));
        Sheet sheet = null;
        int i = workbook.getSheetIndex("sheet"); // sheet表名
        sheet = workbook.getSheetAt(0);
        for (int j = 0; j < sheet.getLastRowNum() + 1; j++) {// getLastRowNum
            // 获取最后一行的行标
            Row row = sheet.getRow(j);
            if (row != null) {
                for (int k = 0; k < row.getLastCellNum(); k++) {// getLastCellNum
                    // 是获取最后一个不为空的列是第几个
                    if (row.getCell(k) != null) { // getCell 获取单元格数据
                        System.out.print(row.getCell(k) + "\t");
                    } else {
                        System.out.print("\t");
                    }
                }
            }
            System.out.println("");
        }
    }
}
