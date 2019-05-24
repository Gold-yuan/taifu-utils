package sichuan.ytf.main.util;

public class AqiUtil {
    
    public static void main(String[] args) {
        int no2 = getAqi(54.0,"no2");
        int o3 = getAqi(59.0,"o3");
        int pm25 = getAqi(78.0,"pm25");
        int so2 = getAqi(34.0,"so2");
        int pm10 = getAqi(193.0,"pm10");
        int co = getAqi(1.6,"co");
        
        System.out.println("no2 :"+no2);
        System.out.println("o3  :"+o3);
        System.out.println("pm25:"+pm25);
        System.out.println("so2 :"+so2);
        System.out.println("pm10:"+pm10);
        System.out.println("co  :"+co);
        
    }
    public static int getAqi(Double value, String pullutionName) {
        
        pullutionName = pullutionName.toLowerCase();
        double aqi = 0.0;
        switch (pullutionName) {
            case "co":
                aqi = getCOIaqi(value);
                break;
            case "no2":
                aqi = getNO2Iaqi(value);
                break;
            case "o3":
                aqi = getO3Iaqi(value);
                break;
            case "pm10":
                aqi = getPM10Iaqi(value);
                break;
            case "pm25":
                aqi = getPM25Iaqi(value);
                break;
            case "so2":
                aqi = getSO2Iaqi(value);
                break;
            default:
        }
        return (int) Math.round(aqi);
    }
    /**
     * 根据污染物浓度，获取空气质量指数类别。
     * 
     * @param value 污染物浓度
     * @param pullutionName 污染物项目
     * @return 空气质量指数类别，如优、良、轻度污染等
     */
    public static String getAqiPhrase(Double value, String pullutionName) {
        if (value < 0) {
            return "污染物浓度值错误";
        }

        pullutionName = pullutionName.toLowerCase();
        double aqi = 0.0;
        switch (pullutionName) {
            case "co":
                aqi = getCOIaqi(value);
                break;
            case "no2":
                aqi = getNO2Iaqi(value);
                break;
            case "o3":
                aqi = getO3Iaqi(value);
                break;
            case "pm10":
                aqi = getPM10Iaqi(value);
                break;
            case "pm25":
                aqi = getPM25Iaqi(value);
                break;
            case "so2":
                aqi = getSO2Iaqi(value);
                break;
            default:
                return "无此污染物";
        }
        return toAqiPhrase(aqi);
    }

    private static String toAqiPhrase(Double value) {
        if (value == null)
            return "暂无数据";
        if (value < 51) {
            return "优";
        } else if (value < 101) {
            return "良";
        } else if (value < 151) {
            return "轻度污染";
        } else if (value < 201) {
            return "中度污染";
        } else if (value < 301) {
            return "重度污染";
        }
        return "严重污染";
    }

    private static Double getPM25Iaqi(double pm25) {
        double aqi = 0;
        double IAQI_H = 0;
        double IAQI_L = 0;
        double BP_H = 0;
        double BP_L = 0;
        int limit = 0;
        if ((limit = 35) == -1 || (BP_H = limit) == -1 || pm25 < limit) {
            IAQI_H = 50;
            IAQI_L = 0;
        } else if ((limit = 75) == -1 || (BP_L = BP_H) == -1 || (BP_H = limit) == -1
                || pm25 < limit) {
            IAQI_H = 100;
            IAQI_L = 50;
        } else if ((limit = 115) == -1 || (BP_L = BP_H) == -1 || (BP_H = limit) == -1
                || pm25 < limit) {
            IAQI_H = 150;
            IAQI_L = 100;
        } else if ((limit = 150) == -1 || (BP_L = BP_H) == -1 || (BP_H = limit) == -1
                || pm25 < limit) {
            IAQI_H = 200;
            IAQI_L = 150;
        } else if ((limit = 250) == -1 || (BP_L = BP_H) == -1 || (BP_H = limit) == -1
                || pm25 < limit) {
            IAQI_H = 300;
            IAQI_L = 200;
        } else if ((limit = 350) == -1 || (BP_L = BP_H) == -1 || (BP_H = limit) == -1
                || pm25 < limit) {
            IAQI_H = 400;
            IAQI_L = 300;
        } else if ((limit = 500) == -1 || (BP_L = BP_H) == -1 || (BP_H = limit) == -1
                || pm25 < limit) {
            IAQI_H = 500;
            IAQI_L = 400;
        }
        aqi = (IAQI_H - IAQI_L) / (BP_H - BP_L) * (pm25 - BP_L) + IAQI_L;
        return aqi;
    }

    private static Double getPM10Iaqi(double pm10) {
        double aqi = 0;
        double IAQI_H = 0;
        double IAQI_L = 0;
        double BP_H = 0;
        double BP_L = 0;
        int limit = 0;
        if ((limit = 50) == -1 || (BP_H = limit) == -1 || pm10 < limit) {
            IAQI_H = 50;
            IAQI_L = 0;
        } else if ((limit = 150) == -1 || (BP_L = BP_H) == -1 || (BP_H = limit) == -1
                || pm10 < limit) {
            IAQI_H = 100;
            IAQI_L = 50;
        } else if ((limit = 250) == -1 || (BP_L = BP_H) == -1 || (BP_H = limit) == -1
                || pm10 < limit) {
            IAQI_H = 150;
            IAQI_L = 100;
        } else if ((limit = 350) == -1 || (BP_L = BP_H) == -1 || (BP_H = limit) == -1
                || pm10 < limit) {
            IAQI_H = 200;
            IAQI_L = 150;
        } else if ((limit = 420) == -1 || (BP_L = BP_H) == -1 || (BP_H = limit) == -1
                || pm10 < limit) {
            IAQI_H = 300;
            IAQI_L = 200;
        } else if ((limit = 500) == -1 || (BP_L = BP_H) == -1 || (BP_H = limit) == -1
                || pm10 < limit) {
            IAQI_H = 400;
            IAQI_L = 300;
        } else if ((limit = 600) == -1 || (BP_L = BP_H) == -1 || (BP_H = limit) == -1
                || pm10 < limit) {
            IAQI_H = 500;
            IAQI_L = 400;
        }
        aqi = (IAQI_H - IAQI_L) / (BP_H - BP_L) * (pm10 - BP_L) + IAQI_L;
        return aqi;
    }

    private static Double getNO2Iaqi(double no2) {
        double aqi = 0;
        double IAQI_H = 0;
        double IAQI_L = 0;
        double BP_H = 0;
        double BP_L = 0;
        int limit = 0;
        if ((limit = 40) == -1 || (BP_H = limit) == -1 || no2 < limit) {
            IAQI_H = 50;
            IAQI_L = 0;
        } else if ((limit = 80) == -1 || (BP_L = BP_H) == -1 || (BP_H = limit) == -1
                || no2 < limit) {
            IAQI_H = 100;
            IAQI_L = 50;
        } else if ((limit = 180) == -1 || (BP_L = BP_H) == -1 || (BP_H = limit) == -1
                || no2 < limit) {
            IAQI_H = 150;
            IAQI_L = 100;
        } else if ((limit = 280) == -1 || (BP_L = BP_H) == -1 || (BP_H = limit) == -1
                || no2 < limit) {
            IAQI_H = 200;
            IAQI_L = 150;
        } else if ((limit = 565) == -1 || (BP_L = BP_H) == -1 || (BP_H = limit) == -1
                || no2 < limit) {
            IAQI_H = 300;
            IAQI_L = 200;
        } else if ((limit = 750) == -1 || (BP_L = BP_H) == -1 || (BP_H = limit) == -1
                || no2 < limit) {
            IAQI_H = 400;
            IAQI_L = 300;
        } else if ((limit = 940) == -1 || (BP_L = BP_H) == -1 || (BP_H = limit) == -1
                || no2 < limit) {
            IAQI_H = 500;
            IAQI_L = 400;
        }
        aqi = (IAQI_H - IAQI_L) / (BP_H - BP_L) * (no2 - BP_L) + IAQI_L;
        return aqi;
    }

    private static Double getCOIaqi(double co) {
        double aqi = 0;
        double IAQI_H = 0;
        double IAQI_L = 0;
        double BP_H = 0;
        double BP_L = 0;
        int limit = 0;
        if ((limit = 2) == -1 || (BP_H = limit) == -1 || co < limit) {
            IAQI_H = 50;
            IAQI_L = 0;
        } else if ((limit = 4) == -1 || (BP_L = BP_H) == -1 || (BP_H = limit) == -1 || co < limit) {
            IAQI_H = 100;
            IAQI_L = 50;
        } else if ((limit = 14) == -1 || (BP_L = BP_H) == -1 || (BP_H = limit) == -1
                || co < limit) {
            IAQI_H = 150;
            IAQI_L = 100;
        } else if ((limit = 24) == -1 || (BP_L = BP_H) == -1 || (BP_H = limit) == -1
                || co < limit) {
            IAQI_H = 200;
            IAQI_L = 150;
        } else if ((limit = 36) == -1 || (BP_L = BP_H) == -1 || (BP_H = limit) == -1
                || co < limit) {
            IAQI_H = 300;
            IAQI_L = 200;
        } else if ((limit = 48) == -1 || (BP_L = BP_H) == -1 || (BP_H = limit) == -1
                || co < limit) {
            IAQI_H = 400;
            IAQI_L = 300;
        } else if ((limit = 60) == -1 || (BP_L = BP_H) == -1 || (BP_H = limit) == -1
                || co < limit) {
            IAQI_H = 500;
            IAQI_L = 400;
        }
        aqi = (IAQI_H - IAQI_L) / (BP_H - BP_L) * (co - BP_L) + IAQI_L;
        return aqi;
    }

    private static Double getO3Iaqi(double o3) {
        double aqi = 0;
        double IAQI_H = 0;
        double IAQI_L = 0;
        double BP_H = 0;
        double BP_L = 0;
        int limit = 0;
        if ((limit = 100) == -1 || (BP_H = limit) == -1 || o3 < limit) {
            IAQI_H = 50;
            IAQI_L = 0;
        } else if ((limit = 160) == -1 || (BP_L = BP_H) == -1 || (BP_H = limit) == -1
                || o3 < limit) {
            IAQI_H = 100;
            IAQI_L = 50;
        } else if ((limit = 215) == -1 || (BP_L = BP_H) == -1 || (BP_H = limit) == -1
                || o3 < limit) {
            IAQI_H = 150;
            IAQI_L = 100;
        } else if ((limit = 265) == -1 || (BP_L = BP_H) == -1 || (BP_H = limit) == -1
                || o3 < limit) {
            IAQI_H = 200;
            IAQI_L = 150;
        } else if ((limit = 800) == -1 || (BP_L = BP_H) == -1 || (BP_H = limit) == -1
                || o3 < limit) {
            IAQI_H = 300;
            IAQI_L = 200;
        } else if ((limit = 1000) == -1 || (BP_L = BP_H) == -1 || (BP_H = limit) == -1
                || o3 < limit) {
            IAQI_H = 400;
            IAQI_L = 300;
        } else if ((limit = 1200) == -1 || (BP_L = BP_H) == -1 || (BP_H = limit) == -1
                || o3 < limit) {
            IAQI_H = 500;
            IAQI_L = 400;
        }
        aqi = (IAQI_H - IAQI_L) / (BP_H - BP_L) * (o3 - BP_L) + IAQI_L;
        return aqi;
    }

    private static Double getSO2Iaqi(double so2) {
        double aqi = 0;
        double IAQI_H = 0;
        double IAQI_L = 0;
        double BP_H = 0;
        double BP_L = 0;
        int limit = 0;
        if ((limit = 50) == -1 || (BP_H = limit) == -1 || so2 < limit) {
            IAQI_H = 50;
            IAQI_L = 0;
        } else if ((limit = 150) == -1 || (BP_L = BP_H) == -1 || (BP_H = limit) == -1
                || so2 < limit) {
            IAQI_H = 100;
            IAQI_L = 50;
        } else if ((limit = 475) == -1 || (BP_L = BP_H) == -1 || (BP_H = limit) == -1
                || so2 < limit) {
            IAQI_H = 150;
            IAQI_L = 100;
        } else if ((limit = 800) == -1 || (BP_L = BP_H) == -1 || (BP_H = limit) == -1
                || so2 < limit) {
            IAQI_H = 200;
            IAQI_L = 150;
        } else if ((limit = 1600) == -1 || (BP_L = BP_H) == -1 || (BP_H = limit) == -1
                || so2 < limit) {
            IAQI_H = 300;
            IAQI_L = 200;
        } else if ((limit = 2100) == -1 || (BP_L = BP_H) == -1 || (BP_H = limit) == -1
                || so2 < limit) {
            IAQI_H = 400;
            IAQI_L = 300;
        } else if ((limit = 2620) == -1 || (BP_L = BP_H) == -1 || (BP_H = limit) == -1
                || so2 < limit) {
            IAQI_H = 500;
            IAQI_L = 400;
        }
        aqi = (IAQI_H - IAQI_L) / (BP_H - BP_L) * (so2 - BP_L) + IAQI_L;
        return aqi;
    }
}

// end
