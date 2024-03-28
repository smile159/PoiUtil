package cn.smile;

import org.mybatis.spring.annotation.MapperScan;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
@MapperScan("cn.smile.mapper")
public class PoiUtilApplication {

    public static void main(String[] args) {
        SpringApplication.run(PoiUtilApplication.class, args);
    }

}
