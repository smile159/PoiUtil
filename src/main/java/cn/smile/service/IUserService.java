package cn.smile.service;

import cn.smile.entity.User;
import com.baomidou.mybatisplus.extension.service.IService;

import java.util.List;

/**
 * <p>
 * 用户表 服务类
 * </p>
 *
 * @author author
 * @since 2024-03-28
 */
public interface IUserService extends IService<User> {


    public List<User> getUserList();
}
