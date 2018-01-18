package cn.hydralisk.common.server;

import cn.hydralisk.common.utils.CommonUtils;
import org.apache.http.entity.ContentType;
import org.apache.http.entity.StringEntity;

/**
 * 表单形式的HttpPost
 * created by yangyebo 2017-12-12 上午10:58
 */
public class FormHttpPostGenerator extends HttpPostGenerator{

    @Override
    protected String getContentType(){
        return ContentType.APPLICATION_FORM_URLENCODED.getMimeType();
    }

    @Override
    protected StringEntity generateStringEntity(HttpRequest httpRequest){
        return new StringEntity(CommonUtils.constitutePostData(
                httpRequest.getRequestParams()), httpRequest.getCharset());
    }
}
