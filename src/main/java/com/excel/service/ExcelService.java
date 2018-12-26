package com.excel.service;


import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;

public interface ExcelService {

    boolean batchImport(String fileName, MultipartFile file) throws Exception;

    boolean exportExcel(HttpServletResponse response) throws Exception;
}
