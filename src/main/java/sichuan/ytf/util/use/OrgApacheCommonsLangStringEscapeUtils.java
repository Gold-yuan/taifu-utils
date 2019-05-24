package sichuan.ytf.util.use;

import org.apache.commons.text.StringEscapeUtils;
import org.junit.Test;

/**
 * @author tfyuan
 *
 */
public class OrgApacheCommonsLangStringEscapeUtils {

	@Test
	public void useStringEscapeUtils() {

		String s3 = "sdf\t\nsddl\t\ndd";
		String s2 = StringEscapeUtils.unescapeJava(s3);
		String s = StringEscapeUtils.escapeJava(s3);
		System.out.println(s2);
		System.out.println(s3);
		System.out.println(s);
	}
}
