/**
 * 
 */
package sichuan.ytf.main.util;

import static java.nio.charset.StandardCharsets.UTF_8;

import org.springframework.util.Base64Utils;

public class Base64Util {

    public static String encodeToStr(byte[] data) {
        return new String(Base64Utils.encode(data), UTF_8);
    }

    public static String encodeToStr(String data) {
        return new String(Base64Utils.encode(data.getBytes(UTF_8)), UTF_8);
    }

    public static byte[] encode(byte[] data) {
        return Base64Utils.encode(data);
    }

    public static byte[] encode(String data) {
        return Base64Utils.encode(data.getBytes(UTF_8));
    }

    public static String decodeToStr(String data) {
        return new String(Base64Utils.decode(data.getBytes(UTF_8)), UTF_8);
    }

    public static String decodeToStr(byte[] data) {
        return new String(Base64Utils.decode(data), UTF_8);
    }

    public static byte[] decode(byte[] data) {
        return Base64Utils.decode(data);
    }

    public static byte[] decode(String data) {
        return Base64Utils.decode(data.getBytes(UTF_8));
    }
}