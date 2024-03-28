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
    @Excel("用户名")
    private String userName;

    /**
     * 性别
     */
    @Excel("性别")
    private Integer sex;

    /**
     * 密码
     */
    @Excel("密码")
    private String password;

    /**
     * 邮箱
     */
    @Excel("邮箱")
    private String email;

    /**
     * 电话号码
     */
    @Excel("电话号码")
    private String phoneNumber;

    /**
     * 状态
     */
    @Excel("状态")
    private String status;

    /**
     * 创建时间
     */
    @Excel("创建时间")
    @TableField(fill = FieldFill.INSERT)
    private LocalDateTime createTime;

    /**
     * 更新时间
     */
    @Excel("更新时间")
    @TableField(fill = FieldFill.INSERT_UPDATE)
    private LocalDateTime updateTime;

}
