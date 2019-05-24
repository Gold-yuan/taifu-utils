package sichuan.ytf.main.util;

import java.io.InputStream;

import org.apache.http.HttpResponse;
import org.apache.http.HttpStatus;
import org.apache.http.client.HttpClient;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.impl.client.DefaultHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.impl.conn.PoolingClientConnectionManager;
import org.apache.http.params.HttpConnectionParams;
import org.apache.http.params.HttpParams;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

@SuppressWarnings("deprecation")
public class HttpClientReadStream {
	private static Logger log = LoggerFactory.getLogger(HttpClientReadStream.class);

	public static void main(String[] args) {
		String proxyUrl = "http://localhost:10092/turbine.stream?delay=1000";
		log.info("\n\nProxy opening connection to: " + proxyUrl + "\n\n");

		HttpGet httpget = null;
		InputStream is = null;
		try {
			httpget = new HttpGet(proxyUrl);

			HttpClient client = HttpClients.createDefault();
			HttpResponse httpResponse = client.execute(httpget);
			int statusCode = httpResponse.getStatusLine().getStatusCode();
			if (statusCode == HttpStatus.SC_OK) {
				is = httpResponse.getEntity().getContent();

				int b = -1;
				byte[] buff = new byte[8];
				while ((b = is.read(buff)) != -1) {
					String data = new String(buff, 0, b);
					System.out.println(data);
				}
			}
		} catch (Exception ex) {
			log.error("Error proxying request: " + proxyUrl, ex);
		} finally {
			if (httpget != null) {
				try {
					httpget.abort();
				} catch (Exception ex) {
					log.error("failed aborting proxy connection.", ex);
				}
			}

			// httpget.abort() MUST be called first otherwise is.close() hangs
			// (because data is still streaming?)
			if (is != null) {
				// this should already be closed by httpget.abort() above
				try {
					is.close();
				} catch (Exception ex) {
					// ignore errors on close
				}
			}
		}

	}

	@SuppressWarnings({ "unused" })
	private static class ProxyConnectionManager {

		private final static PoolingClientConnectionManager threadSafeConnectionManager = new PoolingClientConnectionManager();

		private final static HttpClient httpClient = new DefaultHttpClient(threadSafeConnectionManager);

		static {
			log.debug("Initialize ProxyConnectionManager");
			/* common settings */
			HttpParams httpParams = httpClient.getParams();
			HttpConnectionParams.setConnectionTimeout(httpParams, 5000);
			HttpConnectionParams.setSoTimeout(httpParams, 10000);

			/* number of connections to allow */
			threadSafeConnectionManager.setDefaultMaxPerRoute(400);
			threadSafeConnectionManager.setMaxTotal(400);
		}

	}
}
