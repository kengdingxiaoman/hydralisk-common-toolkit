package cn.hydralisk.common.server;

import org.apache.http.conn.ssl.TrustSelfSignedStrategy;
import org.apache.http.ssl.SSLContexts;

import javax.net.ssl.SSLContext;
import javax.net.ssl.TrustManager;
import javax.net.ssl.X509TrustManager;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.security.KeyManagementException;
import java.security.KeyStore;
import java.security.KeyStoreException;
import java.security.NoSuchAlgorithmException;
import java.security.cert.CertificateException;

/**
 * created by yangyebo 2017-12-12 下午12:28
 */
public class SSLContextFactory{

    /**
     * 绕过验证
     * @return
     * @throws NoSuchAlgorithmException
     * @throws KeyManagementException
     */
    public static SSLContext createIgnoreVerify() throws NoSuchAlgorithmException, KeyManagementException{
        SSLContext sslContext = SSLContext.getInstance("SSLv3");

        // 实现一个X509TrustManager接口，用于绕过验证
        X509TrustManager trustManager = new X509TrustManager() {
            @Override
            public void checkClientTrusted(
                    java.security.cert.X509Certificate[] paramArrayOfX509Certificate,
                    String paramString) throws CertificateException{
                //不校验服务器端证书，什么都不做，视为通过检查
            }

            @Override
            public void checkServerTrusted(
                    java.security.cert.X509Certificate[] paramArrayOfX509Certificate,
                    String paramString) throws CertificateException {
                //不校验服务器端证书，什么都不做，视为通过检查
            }

            @Override
            public java.security.cert.X509Certificate[] getAcceptedIssuers() {
                return null;
            }
        };

        sslContext.init(null, new TrustManager[] { trustManager }, null);
        return sslContext;
    }

    /**
     * 生成证书
     * @param keyStorePath
     * @param keyStorePassWord
     * @return
     */
    public static SSLContext create(String keyStorePath, String keyStorePassWord){
        SSLContext sslContext = null;
        FileInputStream inputStream = null;
        try {
            KeyStore trustStore = KeyStore.getInstance(KeyStore.getDefaultType());
            inputStream = new FileInputStream(new File(keyStorePath));
            trustStore.load(inputStream, keyStorePassWord.toCharArray());
            sslContext = SSLContexts.custom().loadTrustMaterial(trustStore, new TrustSelfSignedStrategy()).build();
        } catch (KeyStoreException | NoSuchAlgorithmException|
                CertificateException | IOException | KeyManagementException e) {
            e.printStackTrace();
        } finally {
            try {
                inputStream.close();
            } catch (IOException e) {
            }
        }
        return sslContext;
    }
}
