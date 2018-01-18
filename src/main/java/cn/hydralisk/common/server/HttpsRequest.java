package cn.hydralisk.common.server;

/**
 * https请求参数
 * created by yangyebo 2017-12-12 下午12:34
 */
public class HttpsRequest extends HttpRequest{

    private String keyStorePath; //密钥存储路径

    private String keyStorePassword; //密钥访问密码

    //getter and setter
    public HttpsRequest(String postUrl) {
        super(postUrl);
    }

    public String getKeyStorePath(){
        return keyStorePath;
    }

    public void setKeyStorePath(String keyStorePath){
        this.keyStorePath = keyStorePath;
    }

    public String getKeyStorePassword(){
        return keyStorePassword;
    }

    public void setKeyStorePassword(String keyStorePassword){
        this.keyStorePassword = keyStorePassword;
    }
}
