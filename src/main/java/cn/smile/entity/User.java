package cn.smile.entity;

import cn.smile.annotation.Excel;
import com.baomidou.mybatisplus.annotation.TableName;
import com.baomidou.mybatisplus.annotation.IdType;
import com.baomidou.mybatisplus.annotation.TableId;

import java.time.LocalDateTime;

import com.baomidou.mybatisplus.annotation.FieldFill;
import com.baomidou.mybatisplus.annotation.TableField;

import java.io.Serializable;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.experimental.Accessors;

/**
 * <p>
 * 用户表
 * </p>
 *
 * @author author
 * @since 2024-03-28
 */
@Data
@AllArgsConstructor
@EqualsAndHashCode(callSuper = false)
@Accessors(chain = true)
@TableName("user")
public class User implements Serializable {

    private static final long serialVersionUID = 1L;

    /**
     * 主键
     */
    @TableId(value = "id", type = IdType.AUTO)
    private Integer id;

    /**
     * 用户名称
     */
    @Excel(value = "用户名", sort = 8)
    private String userName;

    /**
     * 性别
     */
    @Excel(value = "性别", sort = 7)
    private Integer sex;

    /**
     * 密码
     */
    @Excel(value = "密码", sort = 6)
    private String password;

    /**
     * 邮箱
     */
    @Excel(value = "邮箱", sort = 5)
    private String email;

    /**
     * 电话号码
     */
    @Excel(value = "电话号码", sort = 4)
    private String phoneNumber;

    /**
     * 状态
     */
    @Excel(value = "状态", sort = 3)
    private String status;

    /**
     * 创建时间
     */
    @Excel(value = "创建时间", sort = 2)
    @TableField(fill = FieldFill.INSERT)
    private LocalDateTime createTime;

    /**
     * 更新时间
     */
    @Excel(value = "更新时间", sort = 1)
    @TableField(fill = FieldFill.INSERT_UPDATE)
    private LocalDateTime updateTime;

}
