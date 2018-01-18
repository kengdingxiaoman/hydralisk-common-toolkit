package cn.hydralisk.common.server;

import org.apache.http.client.config.RequestConfig;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.entity.StringEntity;
import org.apache.http.protocol.HTTP;

/**
 * 生成 HttpPost
 * created by yangyebo 2017-12-12 上午11:03
 */
public abstract class HttpPostGenerator{

    public HttpPost generate(HttpRequest httpRequest) {
        HttpPost httpPost = new HttpPost(httpRequest.getPostUrl());

        RequestConfig requestConfig = RequestConfig.custom()
                .setConnectTimeout(httpRequest.getTimeout()).build();

        httpPost.setConfig(requestConfig);

        httpPost.setEntity(generateStringEntity(httpRequest));

        httpPost.addHeader(HTTP.CONTENT_TYPE, getContentType());

        return httpPost;
    }

    protected abstract String getContentType();

    protected abstract StringEntity generateStringEntity(HttpRequest httpRequest);
}