package com.excel.mapper;

import com.excel.model.User;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Select;

import java.util.List;

@Mapper
public interface UserMapper {

    void addUser(User sysUser);

    int updateUserByName(User sysUser);

    int selectByName(String name);

    @Select("select * from user")
    List<User> getAllUsers();
}