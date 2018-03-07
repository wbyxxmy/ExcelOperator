package com.yexj.excelOperator;

import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.util.List;

/**
 * Created by xinjian.ye on 2018/3/7.
 */
public interface IFileService {
    /**
     * POI方式导出为Excel
     * @param dataInfoList
     * @return
     * @throws Exception
     */
    SXSSFWorkbook excelExport(List<?> dataInfoList) throws Exception;


    void excelRead() throws Exception;
}
