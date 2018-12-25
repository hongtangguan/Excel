package com.excel.mapper;

import com.excel.model.User;
import org.apache.ibatis.annotations.Mapper;

@Mapper
public interface UserMapper {

    void addUser(User sysUser);

    int updateUserByName(User sysUser);

    int selectByName(String name);
}