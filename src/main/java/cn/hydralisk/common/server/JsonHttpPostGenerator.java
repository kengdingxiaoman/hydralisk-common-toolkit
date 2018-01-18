package cn.hydralisk.common.server;

import cn.hydralisk.common.utils.convert.JsonUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.http.entity.ContentType;
import org.apache.http.entity.StringEntity;

/**
 * json形式的HttpPost
 * created by yangyebo 2017-12-12 上午10:57
 */
public class JsonHttpPostGenerator extends HttpPostGenerator{

    @Override
    protected String getContentType(){
        return ContentType.APPLICATION_JSON.getMimeType();
    }

    @Override
    protected StringEntity generateStringEntity(HttpRequest httpRequest){
        String jsonParam = StringUtils.EMPTY;

        try{
            jsonParam = JsonUtils.writeEntiry2JSON(
                    httpRequest.getRequestParams());
        } catch(Exception ex) {

        }

        return new StringEntity(jsonParam, httpRequest.getCharset());
    }
}
