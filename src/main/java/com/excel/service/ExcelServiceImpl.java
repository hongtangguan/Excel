package com.excel.service;


import com.excel.common.MyException;
import com.excel.mapper.UserMapper;
import com.excel.model.User;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.web.multipart.MultipartFile;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

@Service
@Transactional(readOnly = true)
public class ExcelServiceImpl implements ExcelService {

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
        User user;
        //这里从第二行开始, 第一行是标题
        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            //如果某一行是空 , 直接跳过
            if (row == null){
                continue;
            }

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
}
