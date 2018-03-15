package com.yexj.excelOperator.api;

import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.util.List;
import java.util.Map;

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

    /**
     * 读取Excel
     * @param multipart
     * @param propertiesFileName
     * @param kyeName
     * @param sheetIndex
     * @param clazz
     * @return
     * @throws Exception
     */
    List<?> uploadAndRead(MultipartFile multipart, String propertiesFileName, String kyeName, int sheetIndex,
                          Class<?> clazz) throws Exception;

    @Deprecated
    void excelRead(File file, Class<T> clazz) throws Exception;
}
