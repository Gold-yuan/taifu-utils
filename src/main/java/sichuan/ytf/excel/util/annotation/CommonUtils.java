package sichuan.ytf.excel.util.annotation;

import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Random;
import java.util.UUID;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.lang3.StringUtils;

public class CommonUtils {
    /*********************** string 空判断 start *******************************/

    /** 为空、null、undefined、NaN 返回true */
    public static boolean isBlank(String s) {
        if (StringUtils.isBlank(s)) {
            return true;
        }
        s = s.trim();
        if ("null".equalsIgnoreCase(s)) {
            return true;
        }
        if ("undefined".equalsIgnoreCase(s)) {
            return true;
        }
        if ("NaN".equalsIgnoreCase(s)) {
            return true;
        }
        return false;
    }

    /** 不（为空、null、undefined、NaN 返回true） */
    public static boolean isNotBlank(String s) {
        return !isBlank(s);
    }

    /** 全不为空返回true */
    public static boolean isNoneBlank(String... ss) {
        for (String s : ss) {
            if (isBlank(s)) {
                return false;
            }
        }
        return true;
    }

    /** 任意一个为空返回true */
    public static boolean isAnyBlank(String... ss) {
        for (String s : ss) {
            if (isBlank(s)) {
                return true;
            }
        }
        return false;
    }

    /*********************** string 空判断 end *******************************/

    public static String getRandomStr(int length) {
        if (1 > length || length > 100) {
            return "";
        }
        String str = "";
        while (str.length() < length) {
            str += UUID.randomUUID().toString().replace("-", "");
        }
        return str.substring(0, length);
    }

    /**
     * 0 生成指定长度的随机数
     * 
     * @param length 位数
     * @return
     */
    public static String getRandomNumber(int length) {
        if (1 > length || length > 20) {
            return "";
        }
        String s = "";
        List<Integer> n = new ArrayList<Integer>();
        for (int i = 0; i < 10; i++) {
            n.add(i);
        }
        Random r = new Random();
        for (int i = 0; i < length; i++) {
            s += n.get(r.nextInt(10));
        }
        return s;
    }

    /*********************** 毫秒转换为常见时间 start *******************************/
    private static DecimalFormat decimalFormat = new DecimalFormat("#.#");

    public static String longToTimeStr(long time) {
        double s = toSeconds(time);
        double min = toMinutes(time);
        double hour = toHours(time);
        double day = toDays(time);
        double month = toMonths(time);
        double year = toYears(time);

        String b = "";
        if (year > 1) {
            b += decimalFormat.format(year) + "年，";
        }
        if (month > 1) {
            b += decimalFormat.format(month) + "月，";
        }
        if (day > 1) {
            b += decimalFormat.format(day) + "日，";
        }
        if (hour > 1) {
            b += decimalFormat.format(hour) + "时，";
        }
        if (min > 1) {
            b += decimalFormat.format(min) + "分，";
        }
        b += s + "秒";
        return b;
    }

    private static double toSeconds(long date) {
        return date / 1000.0;
    }

    private static double toMinutes(long date) {
        return toSeconds(date) / 60L;
    }

    private static double toHours(long date) {
        return toMinutes(date) / 60L;
    }

    private static double toDays(long date) {
        return toHours(date) / 24L;
    }

    private static double toMonths(long date) {
        return toDays(date) / 30L;
    }

    private static double toYears(long date) {
        return toMonths(date) / 12L;
    }

    /*********************** 毫秒转换为常见时间 end *******************************/

    /*********************** 地址编码 start *******************************/
    public static String urlEncodeUtf8(String url) {
        return urlEncode(url, StandardCharsets.UTF_8.name());
    }

    public static String urlEncodeGbk(String url) {
        return urlEncode(url, "GBK");
    }

    public static String urlEncode(String url, String enc) {
        if (url == null || url.trim().length() == 0) {
            return "";
        }
        // 判断是否有问号
        int index = url.indexOf("?");
        if (index != -1) {
            // 根据&分组
            String[] params = url.substring(index + 1).split("&");
            for (String param : params) {
                // 找到第一个=号，对后面的进行编码
                int pos = param.indexOf("=");
                if (pos == -1) {
                    continue;
                }
                String value = param.substring(pos + 1);
                try {
                    String encodeValue = URLEncoder.encode(value, enc);
                    url = url.replace(value, encodeValue);
                } catch (UnsupportedEncodingException e) {
                    e.printStackTrace();
                }
            }
        }
        return url;
    }

    /*********************** 地址编码 end *******************************/

    /*********************** 正则匹配字符串 start *******************************/
    /**
     * 获取正则匹配的第一个字符串
     * 
     * @param str      源
     * @param preReg   前置表达式
     * @param regStr   匹配表达式
     * @param afterReg 后置表达式
     * @return
     */
    public static String reg(String str, String preReg, String regStr, String afterReg) {
        Pattern p = Pattern.compile("(?<=" + preReg + ")" + regStr + "(?=" + afterReg + ")");
        Matcher matcher = p.matcher(str);
        while (matcher.find()) {
            return matcher.group();
        }
        return null;
    }

    /**
     * 获取正则匹配的第一个字符串
     * 
     * @param str    源
     * @param regStr 匹配表达式
     * @return
     */
    public static String reg(String str, String regStr) {
        Pattern p = Pattern.compile(regStr);
        Matcher matcher = p.matcher(str);
        while (matcher.find()) {
            return matcher.group();
        }
        return null;
    }

    /**
     * 获取正则匹配的第一个字符串
     * 
     * @param str    源
     * @param preReg 前置表达式
     * @param regStr 匹配表达式
     * @return
     */
    public static String reg(String str, String preReg, String regStr) {
        Pattern p = Pattern.compile("(?<=" + preReg + ")" + regStr);
        Matcher matcher = p.matcher(str);
        while (matcher.find()) {
            return matcher.group();
        }
        return null;
    }

    /**
     * 获取正则匹配的最后一个字符串
     * 
     * @param str      源
     * @param preReg   前置表达式
     * @param regStr   匹配表达式
     * @param afterReg 后置表达式
     * @return
     */
    public static String regLast(String str, String preReg, String regStr, String afterReg) {
        Pattern p = Pattern.compile("(?<=" + preReg + ")" + regStr + "(?=" + afterReg + ")");
        Matcher matcher = p.matcher(str);
        String r = null;
        while (matcher.find()) {
            r = matcher.group();
        }
        return r;
    }

    /**
     * 获取正则匹配的第一个字符串
     * 
     * @param str      源
     * @param regStr   匹配表达式
     * @param afterReg 后置表达式
     * @return
     */
    public static String regAfter(String str, String regStr, String afterReg) {
        Pattern p = Pattern.compile(regStr + "(?=" + afterReg + ")");
        Matcher matcher = p.matcher(str);
        while (matcher.find()) {
            return matcher.group();
        }
        return null;
    }

    /*********************** 正则匹配字符串 end *******************************/

    /*********************** 日期 start *******************************/
    /**
     * 格式化string为Date
     * 
     * @param datestr
     * @return date
     * @throws Exception
     */
    public static Date dateParse(String datestr, String fmtstr) {
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

    public static String dateFormat(Date date, String fmtstr) {
        SimpleDateFormat sdf = new SimpleDateFormat(fmtstr);
        return sdf.format(date);
    }

    public static String dateToYYYYMMdd(Date date) {
        return new SimpleDateFormat("yyyy-MM-dd").format(date);
    }

    public static String dateToYYYYMMddHHmmss(Date date) {
        return new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(date);
    }
    /*********************** 日期 end *******************************/
}
