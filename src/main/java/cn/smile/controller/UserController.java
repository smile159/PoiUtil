package cn.smile.controller;


import cn.smile.entity.User;
import cn.smile.service.IUserService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;

import org.springframework.web.bind.annotation.RestController;

import java.util.List;

/**
 * <p>
 * 用户表 前端控制器
 * </p>
 *
 * @author author
 * @since 2024-03-28
 */
@RestController
@RequestMapping("/user")
public class UserController {


    @Autowired
    private IUserService userService;


    @GetMapping("list")
    public List<User> getUserList() {
        return userService.getUserList();
    }

}
