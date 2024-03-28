package cn.smile.util;

import javax.servlet.http.HttpServletResponse;
import java.io.UnsupportedEncodingException;

public class ResponseUtil {


    public static void setFileResponseHeader(HttpServletResponse response, String fileName) throws UnsupportedEncodingException {
        response.setContentType("application/octet-stream;charset=UTF-8");
        String urlEncodeName = EncodeUtil.urlEncode(fileName);
        response.setHeader("Content-Disposition", "attachment;filename=" + urlEncodeName);
    }
}
