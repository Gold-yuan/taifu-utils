package sichuan.ytf.main.util;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class DateUtil {

    /**
     * 格式化string为Date
     * 
     * @param datestr
     * @return date
     * @throws Exception
     */
    public static Date parseDate(String datestr, String fmtstr) {
        if (null == datestr || "".equals(datestr)) {
            return null;
        }
        try {
            if (fmtstr == null || fmtstr.trim().length() == 0) {
                if (datestr.indexOf(':') > 0) {
                    fmtstr = "yyyy-MM-dd HH:mm:ss";
                } else {
                    fmtstr = "yyyy-MM-dd";
                }
            }
            SimpleDateFormat sdf = new SimpleDateFormat(fmtstr);
            return sdf.parse(datestr);
        } catch (Exception e) {
            System.out.println("字符串转换日期失败");
            // throw new BusinessException("字符串转换日期失败",e);
        }
        return null;
    }
    
    
    public static String getYYYYMMdd(Date date){
        return new SimpleDateFormat("yyyy-MM-dd").format(date);
    }
    public static String getYYYYMMddHHmmss(Date date){
        return new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(date);
    }
    public static String getYYYYMMddHHmmssSSS(Date date){
        return new SimpleDateFormat("yyyy-MM-dd HH:mm:ss SSS").format(date);
    }
    public static String nowYYYYMMdd(){
        return new SimpleDateFormat("yyyy-MM-dd").format(new Date());
    }
    public static String nowYYYYMMddHHmmss(){
        return new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(new Date());
    }
    public static String nowHHmmssSSS(){
        return new SimpleDateFormat("HH:mm:ss SSS").format(new Date());
    }
    public static Date getYYYYMMdd(String date){
        try {
            return new SimpleDateFormat("yyyy-MM-dd").parse(date);
        } catch (ParseException e) {
            e.printStackTrace();
        }
        return null;
    }
}

// end
