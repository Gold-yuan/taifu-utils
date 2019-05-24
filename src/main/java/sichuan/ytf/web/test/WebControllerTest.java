package sichuan.ytf.web.test;

import static java.nio.charset.StandardCharsets.UTF_8;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.io.FileUtils;
import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import sichuan.ytf.main.util.GsonUtils;

/**
 * 证书相关测试 测试前： 1. 顺序启动CA，RA，确保CA、RA配置文件中MSP_ID一致； 2. 修改下面常量MSP_ID与CA、RA一致； 3.
 * 修改下面常量RA_PROJECT_URL与RA项目访问路径一致
 */
public class WebControllerTest {
	private static final Logger log = LoggerFactory.getLogger(WebControllerTest.class);
	/** 自身 MSP ID */
	private static final String MSP_ID = "org1";
	/** RA项目访问路径 */
	private static final String RA_PROJECT_URL = "localhost:8080";
	/** 测试文件路径 */
	private static final String TEST_RESOURCES_PATH = "src/test/resources";

	/**
	 * 正常测试，测试前确保数据库清空
	 */
	@Test
	public void testNormal() throws Exception {
		final String[] testFileNames = { "test-normal.json" };
		List<TestRequest> testRequests = loadTestRequests(testFileNames);
		log.info("测试请求：\n{}", testRequests);

		HttpTestUtils.sendTestRequests(testRequests);
	}

	/**
	 * 异常测试，测试前确保数据库清空
	 */
	@Test
	public void testException() throws Exception {
		final String[] testFileNames = { "test-exception.json" };
		List<TestRequest> testRequests = loadTestRequests(testFileNames);
		log.info("测试请求：\n{}", testRequests);

		HttpTestUtils.sendTestRequests(testRequests);
	}

	private List<TestRequest> loadTestRequests(String[] testFiles) throws IOException {
		List<TestRequest> testRequests = new ArrayList<>();

		for (String name : testFiles) {
			String result = FileUtils.readFileToString(new File(TEST_RESOURCES_PATH, name), UTF_8);
			result = result.replaceAll("orgcn", MSP_ID);
			result = result.replaceAll("localhost:8080", RA_PROJECT_URL);
			testRequests.addAll(GsonUtils.fromJsonToList(result, TestRequest.class));
		}
		return testRequests;
	}
}
