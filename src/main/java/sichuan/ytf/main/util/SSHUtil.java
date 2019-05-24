/**
 * 
 */
package sichuan.ytf.main.util;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.UnsupportedEncodingException;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import ch.ethz.ssh2.Connection;
import ch.ethz.ssh2.SCPClient;
import ch.ethz.ssh2.Session;
import ch.ethz.ssh2.StreamGobbler;

/**
 * 远程执行linux的shell script
 * 
 * @author Ickes
 * @since V0.1
 */
public class SSHUtil {
	private Logger log = LoggerFactory.getLogger(SSHUtil.class);
	// 字符编码默认是utf-8
	private String defaultChart = "UTF-8";
	private Connection conn;
	private String ip;
	private int port;
	private String userName;
	private String userPwd;

	private SSHUtil(String ip, int port, String userName, String userPwd) {
		this.ip = ip;
		this.userName = userName;
		this.userPwd = userPwd;
		this.port = port;
	}

	public static SSHUtil getInstance(String ip, int port, String userName, String password) {
		SSHUtil util = new SSHUtil(ip, port, userName, password);
		return util;
	}

	public static SSHUtil getInstance(String ip, String userName, String password) {
		return getInstance(ip, 22, userName, password);
	}

	public void addShell(String cmd) {
		String echo = "echo \"" + cmd + "\" >> temp.sh";
		String result = execute(echo);
		log.info(result);
	}

	public String bathExecute() {
		String sh = execute("cat temp.sh");
		log.info("将要执行：{}", sh);
		String shell = "bash temp.sh";

		String result = execute(shell);
		execute("rm -f temp.sh");
		log.info("执行结果：{}", result);
		return result;
	}

	public void getFile(String remoteFile, String path) {
		try {
			if (login()) {
				log.info("认证成功!");
				SCPClient scpClient = conn.createSCPClient();
				scpClient.get(remoteFile, path);
			}
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			conn.close();
		}
	}

	public void putFile(String localFile, String remoteTargetDirectory) throws IOException {
		try {
			if (login()) {
				SCPClient scpClient = conn.createSCPClient();
				scpClient.put(localFile, remoteTargetDirectory);
			}
		} catch (IOException e) {
			throw e;
		}
	}

	/**
	 * 远程登录linux的主机
	 * 
	 * @author Ickes
	 * @since V0.1
	 * @return 登录成功返回true，否则返回false
	 */
	private Boolean login() {
		boolean flg = false;
		try {
			conn = new Connection(ip, port);
			conn.connect();// 连接
			flg = conn.authenticateWithPassword(userName, userPwd);// 认证
		} catch (IOException e) {
			log.info("({}) ({}) ({})", ip, userName, userPwd);
			log.error("ssh认证失败", e);
		}
		return flg;
	}

	/**
	 * @author Ickes 远程执行shll脚本或者命令
	 * @param cmd
	 *            即将执行的命令
	 * @return 命令执行完后返回的结果值
	 * @since V0.1
	 */
	private String execute(String cmd) {
		String result = "";
		try {
			if (login()) {
				Session session = conn.openSession();// 打开一个会话
				session.execCommand(cmd);// 执行命令
				result = processStdout(session.getStdout(), defaultChart);
				// 如果为得到标准输出为空，说明脚本执行出错了
				// if(StringUtils.isBlank(result)){
				result += processStdout(session.getStderr(), defaultChart);
				// }
				conn.close();
				session.close();
			}
		} catch (IOException e) {
			log.error("ssh执行的命令失败", e);
		}
		return result;
	}

	/**
	 * @author Ickes 远程执行shll脚本或者命令
	 * @param cmd
	 *            即将执行的命令
	 * @return 命令执行成功后返回的结果值，如果命令执行失败，返回空字符串，不是null
	 * @since V0.1
	 */
	protected String executeSuccess(String cmd) {
		String result = "";
		try {
			if (login()) {
				Session session = conn.openSession();// 打开一个会话
				session.execCommand(cmd);// 执行命令
				result = processStdout(session.getStdout(), defaultChart);
				conn.close();
				session.close();
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
		return result;
	}

	/**
	 * 解析脚本执行返回的结果集
	 * 
	 * @author Ickes
	 * @param in
	 *            输入流对象
	 * @param charset
	 *            编码
	 * @since V0.1
	 * @return 以纯文本的格式返回
	 */
	private String processStdout(InputStream in, String charset) {
		InputStream stdout = new StreamGobbler(in);
		StringBuffer buffer = new StringBuffer();
		;
		try {
			BufferedReader br = new BufferedReader(new InputStreamReader(stdout, charset));
			String line = null;
			while ((line = br.readLine()) != null) {
				buffer.append(line + "\n");
			}
			br.close();
		} catch (UnsupportedEncodingException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return buffer.toString();
	}

	public void setCharset(String charset) {
		this.defaultChart = charset;
	}

	public void close() {
		try {
			if (conn != null) {
				conn.close();
			}
		} catch (Exception e) {
		}
	}

	public static void main(String[] args) throws IOException {
		// linux ip地址
		String ip = "";
		// linux 用户名
		String userName = "";
		// linux 登录密码
		String password = "";
		// 本地文件路径，到文件名
		String localPath = "";
		// 远程文件夹
		String remotePath = "";
		// 文件名
		String fileName = "";
		
		SSHUtil ssh = SSHUtil.getInstance(ip, userName, password);
		// 上传文件
		ssh.putFile(localPath, remotePath);

		// 执行ssh命令，解压文件
		if (fileName.endsWith(".zip")) {
			ssh.addShell("mkdir -p " + remotePath);
			ssh.addShell("cd " + remotePath);
			ssh.addShell("unzip -oq " + fileName);
			ssh.addShell("rm -f " + fileName);
			ssh.bathExecute();
		}
		ssh.close();
	}
}