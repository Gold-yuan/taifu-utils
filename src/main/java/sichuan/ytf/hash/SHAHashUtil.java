package sichuan.ytf.hash;

import java.io.UnsupportedEncodingException;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;

/**
 * 该算法计算消息摘要 顺序有关
 */
public class SHAHashUtil {

	/**
	 * @param strSrc
	 *            要加密的字符串
	 * @return SHA-256 加密算法的字符串
	 * @throws UnsupportedEncodingException 
	 */
	public static String encryptSHA256(String strSrc) throws UnsupportedEncodingException {
		if (strSrc == null || strSrc.trim().length() == 0) {
			return null;
		}
		return encrypt(strSrc, "SHA-256");
	}

	public static String encryptSHA256(byte[] bt) {
		if (bt == null)
			return null;
		return encryptSHA256(bt, "SHA-256");
	}

	public static String encrypt(byte[] bt, String encName) {
		if (bt == null)
			return null;
		return encryptSHA256(bt, encName);
	}

	/**
	 * 对字符串加密,加密算法使用MD5, SHA-1,SHA-224,SHA-256,SHA-384,SHA-512。
	 * 
	 * 主要适用于数字签名标准（DigitalSignature Standard DSS）里面定义的数字签名算法（Digital Signature
	 * Algorithm DSA）。
	 * 
	 * @param strSrc
	 *            要加密的字符串
	 * @param encName
	 *            加密类型
	 * @return
	 * @throws UnsupportedEncodingException 
	 */
	public static String encrypt(String strSrc, String encName) throws UnsupportedEncodingException {
		return encryptSHA256(strSrc.getBytes("utf-8"), encName);
	}

	private static String encryptSHA256(byte[] bt, String name) {
		MessageDigest md = null;
		String strDes = null;
		try {
			md = MessageDigest.getInstance(name);
			md.update(bt);
			strDes = bytes2Hex(md.digest()); // to HexString
		} catch (NoSuchAlgorithmException e) {
			return null;
		}
		return strDes;
	}

	private static String bytes2Hex(byte[] bts) {
		StringBuffer sb = new StringBuffer();
		for (int i = 0; i < bts.length; i++) {
			String hex = Integer.toHexString(bts[i] & 0xFF);
			if (hex.length() == 1) {
				hex = '0' + hex;
			}
			sb.append(hex.toUpperCase());
		}
		return sb.toString();
	}

	public static void main(String args[]) throws Exception {
		String[] en = { "MD5", "SHA-1", "SHA-224", "SHA-256", "SHA-384", "SHA-512" };

		for (int i = 0; i < en.length; i++) {
			String s2 = encrypt("my name is typhoon", en[i]);
			System.out.println(s2 + " \n" + s2.length() + "位  " + en[i]+"\n");
		}
	}
}
