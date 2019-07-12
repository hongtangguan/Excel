package com.excel.service;


import com.excel.common.MyException;
import com.excel.mapper.UserMapper;
import com.excel.model.User;
import com.excel.utils.ExcelData;
import com.excel.utils.ExportExcelUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

@Service
@Transactional(readOnly = true)
public class ExcelServiceImpl implements ExcelService {

    private Logger logger = LoggerFactory.getLogger(getClass());

    @Autowired
    private UserMapper userMapper;


    @Transactional(readOnly = false,rollbackFor = Exception.class)
    @Override
    public boolean batchImport(String fileName, MultipartFile file) throws Exception {

        boolean notNull = false;
        List<User> userList = new ArrayList<User>();
        if (!fileName.matches("^.+\\.(?i)(xls)$") && !fileName.matches("^.+\\.(?i)(xlsx)$")) {
            throw new MyException("上传文件格式不正确");
        }
        boolean isExcel2003 = true;
        if (fileName.matches("^.+\\.(?i)(xlsx)$")) {
            isExcel2003 = false;
        }
        InputStream is = file.getInputStream();
        Workbook wb = null;
        if (isExcel2003) {
            //wb = new HSSFWorkbook(is);  .xls 格式的不行
            wb = new XSSFWorkbook(is);
        } else {
            wb = new XSSFWorkbook(is);
        }

        Sheet sheet = wb.getSheetAt(0);
        if(sheet!=null){
            notNull = true;
        }

        int lastRowNum = sheet.getLastRowNum();
        System.out.println(lastRowNum+"    ...........................");


        User user;
        //这里从第二行开始, 第一行是标题
        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            //如果某一行是空 , 直接跳过
            if (row == null){
                continue;
            }

            int defaultColumnWidth = sheet.getDefaultColumnWidth();
            System.out.println("defaultColumnWidth...................."+defaultColumnWidth);

            user = new User();

            if( row.getCell(0).getCellType() !=1){
                //throw new MyException("导入失败(第"+(r+1)+"行,姓名请设为文本格式)");
                continue;
            }

            String name = row.getCell(0).getStringCellValue();
            if(StringUtils.isBlank(name)){
                //throw new MyException("导入失败(第"+(r+1)+"行,姓名未填写)");
                continue;
            }

            row.getCell(1).setCellType(Cell.CELL_TYPE_STRING);
            String phone = row.getCell(1).getStringCellValue();
            if(StringUtils.isBlank(phone)){
                //throw new MyException("导入失败(第"+(r+1)+"行,电话未填写)");
                continue;
            }
            String address = row.getCell(2).getStringCellValue();
            if(StringUtils.isBlank(address)){
                throw new MyException("导入失败(第"+(r+1)+"行,不存在此单位或单位未填写)");
            }

            Date date;
            if(row.getCell(3).getCellType() !=0){
                throw new MyException("导入失败(第"+(r+1)+"行,入职日期格式不正确或未填写)");
            }else{
                date = row.getCell(3).getDateCellValue();
            }

            String des = row.getCell(4).getStringCellValue();
            if (StringUtils.isBlank(des)) {
                continue;
            }
            user.setName(name);
            user.setPhone(phone);
            user.setAddress(address);
            user.setEnrolDate(date);
            user.setDes(des);

            userList.add(user);
        }
        for (User userResord : userList) {
            String name = userResord.getName();
            //int cnt = userMapper.selectByName(name);
            int cnt = 0;
            if (cnt == 0) {
                userMapper.addUser(userResord);
                System.out.println(" 插入 "+userResord);
            } else {
                userMapper.updateUserByName(userResord);
                System.out.println(" 更新 "+userResord);
            }
        }
        return notNull;
    }

    @Override
    public boolean exportExcel(HttpServletResponse response) {

        logger.info("导出excel开始...............................");
        ExcelData data = new ExcelData();
        data.setName("users");
        //标题
        List<String> titles = new ArrayList();
        titles.add("姓名");
        titles.add("电话");
        titles.add("地址");
        titles.add("注册日期");
        titles.add("备注");
        data.setTitles(titles);

        List<User> allJobs = userMapper.getAllUsers();
        logger.info("总行数:" + allJobs.size());
        List<List<Object>> rows = new ArrayList();

        for (int i = 0; i < allJobs.size(); i++) {
            List<Object> row = new ArrayList();
            row.add(allJobs.get(i).getName());
            row.add(allJobs.get(i).getPhone());
            row.add(allJobs.get(i).getAddress());
            row.add(allJobs.get(i).getEnrolDate());
            row.add(allJobs.get(i).getDes());
            rows.add(row);
        }
        data.setRows(rows);

        ExportExcelUtils.exportExcel(response, "users.xlsx", data);
        try {
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }

        logger.info("导出完成................");

        return true;

    }
}
