package cn.hydralisk.common.server;

import cn.hydralisk.common.constants.WebConstants;
import org.apache.commons.lang3.StringUtils;
import org.apache.http.client.methods.HttpPost;

import java.io.Serializable;
import java.util.Map;

/**
 * http请求参数的的封装
 * created by yangyebo 2017-12-12 上午9:03
 */
public class HttpRequest implements Serializable{

    private static final long serialVersionUID = -1976145721625827398L;

    private String postUrl; //访问地址

    private Map<String, String> requestParams; //请求参数 key,value

    private String charset; //编码

    private int timeout; //超时时间

    private ContentType contentType;

    public HttpRequest(String postUrl) {
        this.postUrl = postUrl;
    }

    public HttpPost build(){
        if(StringUtils.isBlank(charset)) {
            this.charset = WebConstants.UTF8;
        }
        if(timeout <= 0) {
            this.timeout = WebConstants.DEFAULT_TIME_OUT_MILLISECONDS;
        }

        HttpPostGenerator httpPostGenerator
                = HttpPostGeneratorFactory.chooseGenerator(this);

        return httpPostGenerator.generate(this);
    }

    public HttpRequest requestParams(Map<String, String> requestParams) {
        this.requestParams = requestParams;
        return this;
    }

    public HttpRequest charset(String charset) {
        this.charset = charset;
        return this;
    }

    public HttpRequest timeout(int timeout) {
        this.timeout = timeout;
        return this;
    }

    public static enum ContentType{
        JSON, XML, FORM;

        private ContentType() {
        }
    }

    //getter and setter
    public String getPostUrl(){
        return postUrl;
    }

    public void setPostUrl(String postUrl){
        this.postUrl = postUrl;
    }

    public Map<String, String> getRequestParams(){
        return requestParams;
    }

    public void setRequestParams(Map<String, String> requestParams){
        this.requestParams = requestParams;
    }

    public String getCharset(){
        return charset;
    }

    public void setCharset(String charset){
        this.charset = charset;
    }

    public int getTimeout(){
        return timeout;
    }

    public void setTimeout(int timeout){
        this.timeout = timeout;
    }

    public ContentType getContentType(){
        return contentType;
    }

    public void setContentType(ContentType contentType){
        this.contentType = contentType;
    }
}
