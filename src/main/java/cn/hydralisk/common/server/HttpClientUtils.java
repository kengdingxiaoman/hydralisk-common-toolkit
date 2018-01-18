package cn.hydralisk.common.server;

import cn.hydralisk.common.constants.CharConstants;
import cn.hydralisk.common.constants.WebConstants;
import cn.hydralisk.common.utils.CommonUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.http.HttpEntity;
import org.apache.http.HttpResponse;
import org.apache.http.HttpStatus;
import org.apache.http.client.ClientProtocolException;
import org.apache.http.client.ResponseHandler;
import org.apache.http.client.config.RequestConfig;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.config.Registry;
import org.apache.http.config.RegistryBuilder;
import org.apache.http.conn.socket.ConnectionSocketFactory;
import org.apache.http.conn.socket.PlainConnectionSocketFactory;
import org.apache.http.conn.ssl.SSLConnectionSocketFactory;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.impl.conn.PoolingHttpClientConnectionManager;
import org.apache.http.util.EntityUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.util.Assert;

import javax.net.ssl.SSLContext;
import javax.servlet.http.HttpServletRequest;
import java.io.IOException;
import java.util.Map;

/**
 * httpclient工具类
 * @author master.yang
 * @version $Id: HttpClientUtils.java, v 0.1 2014-6-9 上午10:52:09 Administrator Exp $
 */
public class HttpClientUtils {

    private static Logger logger = LoggerFactory.getLogger(HttpClientUtils.class);

    public static String doGet(String getUrl, Map<String, String> requestParams) {
        logger.debug("Get请求，参数: {}", CommonUtils.showDetails(requestParams));

        CloseableHttpClient httpclient = HttpClients.createDefault();
        try {
            String targetUrl = getUrl + CharConstants.QUESTION_MARK
                    + CommonUtils.constitutePostData(requestParams);

            HttpGet httpget = new HttpGet(targetUrl);

            RequestConfig requestConfig = RequestConfig.custom()
                    .setConnectTimeout(WebConstants.DEFAULT_TIME_OUT_MILLISECONDS).build();
            httpget.setConfig(requestConfig);

            ResponseHandler<String> responseHandler = new ResponseHandler<String>() {
                public String handleResponse(
                        final HttpResponse response) throws IOException{
                    int status = response.getStatusLine().getStatusCode();
                    if (status >= HttpStatus.SC_OK && status < HttpStatus.SC_MULTIPLE_CHOICES) {
                        HttpEntity entity = response.getEntity();
                        return entity != null ? EntityUtils.toString(entity) : null;
                    } else {
                        throw new ClientProtocolException("Unexpected response status: " + status);
                    }
                }
            };
            return httpclient.execute(httpget, responseHandler);
        } catch(Exception ex) {
            logger.error("请求发生发生异常", ex);
            throw new RuntimeException("Get请求发生系统异常");
        } finally {
            try{
                httpclient.close();
            } catch(Exception ex) {
                logger.error("关闭httpclient发生异常", ex);
            }
        }
    }

    public static String doHttpPost(HttpRequest httpRequest) {
        logger.debug("发送请求: {}", CommonUtils.showDetails(httpRequest));

        CloseableHttpClient httpclient = HttpClients.createDefault();
        try {
            HttpPost httpPost = httpRequest.build();

            ResponseHandler<String> responseHandler = new ResponseHandler<String>() {
                public String handleResponse(
                        final HttpResponse response) throws IOException{
                    int status = response.getStatusLine().getStatusCode();
                    if (status >= HttpStatus.SC_OK && status < HttpStatus.SC_MULTIPLE_CHOICES) {
                        HttpEntity entity = response.getEntity();
                        return entity != null ? EntityUtils.toString(entity) : null;
                    } else {
                        throw new ClientProtocolException("Unexpected response status: " + status);
                    }
                }
            };
            return httpclient.execute(httpPost, responseHandler);
        } catch(Exception ex) {
            logger.error("请求发生发生异常", ex);
            throw new RuntimeException("Post请求发生系统异常");
        } finally {
            try{
                httpclient.close();
            } catch(Exception ex) {
                logger.error("关闭httpclient发生异常", ex);
            }
        }
    }

    public static String doHttpsPost(HttpsRequest httpsRequest) {
        logger.info("发送请求: {}", CommonUtils.showDetails(httpsRequest));

        CloseableHttpClient httpclient = null;

        try {
            SSLContext sslcontext = SSLContextFactory.create(httpsRequest.getKeyStorePath(), httpsRequest.getKeyStorePassword());

            Registry<ConnectionSocketFactory> socketFactoryRegistry = RegistryBuilder.<ConnectionSocketFactory>create()
                    .register(WebConstants.HTTP_PROTOCOL_SCHEME, PlainConnectionSocketFactory.INSTANCE)
                    .register(WebConstants.HTTPS_PROTOCOL_SCHEME, new SSLConnectionSocketFactory(sslcontext))
                    .build();
            PoolingHttpClientConnectionManager connManager = new PoolingHttpClientConnectionManager(socketFactoryRegistry);
            HttpClients.custom().setConnectionManager(connManager);

            httpclient = HttpClients.custom().setConnectionManager(connManager).build();

            HttpPost httpPost = httpsRequest.build();

            ResponseHandler<String> responseHandler = new ResponseHandler<String>() {
                public String handleResponse(
                        final HttpResponse response) throws IOException{
                    int status = response.getStatusLine().getStatusCode();
                    if (status >= HttpStatus.SC_OK && status < HttpStatus.SC_MULTIPLE_CHOICES) {
                        HttpEntity entity = response.getEntity();
                        return entity != null ? EntityUtils.toString(entity) : null;
                    } else {
                        throw new ClientProtocolException("Unexpected response status: " + status);
                    }
                }
            };
            return httpclient.execute(httpPost, responseHandler);
        } catch(Exception ex) {
            logger.error("请求发生发生异常", ex);
            throw new RuntimeException("Post请求发生系统异常");
        } finally {
            try{
                httpclient.close();
            } catch(Exception ex) {
                logger.error("关闭httpclient发生异常", ex);
            }
        }
    }

    public static String doHttpsPostIgnoreVerify(HttpRequest httpRequest) {
        logger.info("发送请求: {}", CommonUtils.showDetails(httpRequest));

        CloseableHttpClient httpclient = null;

        try {
            SSLContext sslcontext = SSLContextFactory.createIgnoreVerify();

            Registry<ConnectionSocketFactory> socketFactoryRegistry = RegistryBuilder.<ConnectionSocketFactory>create()
                    .register(WebConstants.HTTP_PROTOCOL_SCHEME, PlainConnectionSocketFactory.INSTANCE)
                    .register(WebConstants.HTTPS_PROTOCOL_SCHEME, new SSLConnectionSocketFactory(sslcontext))
                    .build();
            PoolingHttpClientConnectionManager connManager = new PoolingHttpClientConnectionManager(socketFactoryRegistry);
            HttpClients.custom().setConnectionManager(connManager);

            httpclient = HttpClients.custom().setConnectionManager(connManager).build();

            HttpPost httpPost = httpRequest.build();

            ResponseHandler<String> responseHandler = new ResponseHandler<String>() {
                public String handleResponse(
                        final HttpResponse response) throws IOException{
                    int status = response.getStatusLine().getStatusCode();
                    if (status >= HttpStatus.SC_OK && status < HttpStatus.SC_MULTIPLE_CHOICES) {
                        HttpEntity entity = response.getEntity();
                        return entity != null ? EntityUtils.toString(entity) : null;
                    } else {
                        throw new ClientProtocolException("Unexpected response status: " + status);
                    }
                }
            };
            return httpclient.execute(httpPost, responseHandler);
        } catch(Exception ex) {
            logger.error("请求发生发生异常", ex);
            throw new RuntimeException("Post请求发生系统异常");
        } finally {
            try{
                httpclient.close();
            } catch(Exception ex) {
                logger.error("关闭httpclient发生异常", ex);
            }
        }
    }

    /**
     * 获取请求的IP地址
     * @param request
     * @return
     */
    public static String getRequestIP(HttpServletRequest request) {
        Assert.notNull(request, "request is null");
        String ip = request.getHeader("X-Real-IP");
        if (ipCanBeObtain(ip)) {
            return ip;
        }

        ip = request.getHeader("X-forwarded-for");
        if (ipCanBeObtain(ip)) {
            // 多次反向代理后会有多个IP值，第一个为真实IP。
            int index = ip.indexOf(CharConstants.COMMA);
            return index == -1 ? ip : ip.substring(0, index);
        }

        return request.getRemoteAddr();
    }

    private static boolean ipCanBeObtain(String ip) {
        return StringUtils.isNotBlank(ip) && !"unknown".equalsIgnoreCase(ip);
    }
}
