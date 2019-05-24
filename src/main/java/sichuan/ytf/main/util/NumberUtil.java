package sichuan.ytf.main.util;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.text.DecimalFormat;

public class NumberUtil {

    /**
     * @param d 
     * @param i
     * @return 四舍五入保留i为小数
     */
    private static DecimalFormat df1 = new DecimalFormat("#.0");
    private static DecimalFormat df2 = new DecimalFormat("#.00");
    private static DecimalFormat df3 = new DecimalFormat("#.000");
    
    public static String formatOne(Double d){
        return df1.format(d);
    }
    public static String formatTwo(Double d){
        return df2.format(d);
    }
    public static String formatThree(Double d){
        return df3.format(d);
    }
    
    public Double format(double d, int i){
        BigDecimal db = new BigDecimal(d);
        return db.setScale(i,   BigDecimal.ROUND_HALF_UP).doubleValue();
    }
    public Double format(String d, int i){
        BigDecimal db = new BigDecimal(d);
        return db.setScale(i,   BigDecimal.ROUND_HALF_UP).doubleValue();
    }
    public Double format(int d, int i){
        BigDecimal db = new BigDecimal(d);
        return db.setScale(i,   BigDecimal.ROUND_HALF_UP).doubleValue();
    }
    public Double format(long d, int i){
        BigDecimal db = new BigDecimal(d);
        return db.setScale(i,   BigDecimal.ROUND_HALF_UP).doubleValue();
    }
    public Double format(BigInteger d, int i){
        BigDecimal db = new BigDecimal(d);
        return db.setScale(i,   BigDecimal.ROUND_HALF_UP).doubleValue();
    }
    
}

// end
