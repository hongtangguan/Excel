package com.excel.controller;

import com.excel.service.ExcelService;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;


@Api(value = "/test/", tags = "excel导出导入测试")
@RestController
@RequestMapping("/test/")
public class ExcelController {

    @Autowired
    private ExcelService testService;

    @ApiOperation(value = "导入数据",notes = "导入数据")
    @PostMapping("/import")
    public boolean addUser(@RequestParam("file") MultipartFile file) {
        boolean a = false;
        String fileName = file.getOriginalFilename();
        try {
             a = testService.batchImport(fileName, file);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return  a;
    }



    @RequestMapping(value = "/excel", method = RequestMethod.GET)
    @ApiOperation(value = "导出到excel")//,produces="application/octet-stream"
    public boolean excel(HttpServletResponse response) throws Exception {
        return testService.exportExcel(response);

    }


}
