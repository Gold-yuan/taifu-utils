package sichuan.ytf.main.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.net.URL;
import java.util.Properties;

public class PropertiesUtil {
    private PropertiesUtil() {}

    private static Properties p = new Properties();

    static {
        URL url = PropertiesUtil.class.getResource("/");
        File dir = new File(url.getFile());
        if (dir.isDirectory()) {
            File[] files = dir.listFiles();
            for (File file : files) {
                if (file.getName().endsWith(".properties")) {
                    Properties prop = new Properties();
                    try {
                        prop.load(new FileInputStream(file));
                        p.putAll(prop);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
            }
        }
    }

    public static String getValue(String key) {
        return p.getProperty(key);
    }

    public static void setFilePath(String directory) {
        URL url = PropertiesUtil.class.getResource("/");
        File dir = new File(url.getFile());
        System.out.println(dir.getParent());
        System.out.println(System.getProperty("user.dir"));
        if (dir.isDirectory()) {
            File[] files = dir.listFiles();
            for (File file : files) {
                if (file.getName().endsWith(".properties")) {
                    Properties prop = new Properties();
                    try {
                        prop.load(new FileInputStream(file));
                        p.putAll(prop);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
            }
        }
    }

    private static void init() {
        if (p != null) {
            return;
        }
        System.out.println("初始化...");
        p = new Properties();
        URL url = PropertiesUtil.class.getResource("/");
        File dir = new File(url.getFile());
        if (dir.isDirectory()) {
            File[] files = dir.listFiles();
            for (File file : files) {
                if (file.getName().endsWith(".properties")) {
                    Properties prop = new Properties();
                    try {
                        prop.load(new FileInputStream(file));
                        p.putAll(prop);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
            }
        }
    }

    /**
     * @param fileName classPath为根目录/
     * @return 一个properties对象
     */
    private static Properties getPropertiesFile(String fileName) {
        Properties prop = new Properties();
        try {
            if (fileName.startsWith("/")) {
                prop.load(PropertiesUtil.class.getResourceAsStream(fileName));
            } else {
                prop.load(PropertiesUtil.class.getClassLoader().getResourceAsStream(fileName));
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return prop;
    }

    public static void main(String[] args) {
        getPropertiesFile("log4j.properties");
        init();
        String key1 = PropertiesUtil.getValue("activemq.password");
        String key2 = PropertiesUtil.getValue("driver");
        String key3 = PropertiesUtil.getValue("log4j.rootLogger");
        System.out.println(key1);
        System.out.println(key2);
        System.out.println(key3);
        PropertiesUtil.setFilePath("mqconfig.properties");
    }
}

// end
