package cn.hydralisk.common.server;

/**
 * HttpPostGenerator 工厂类
 * created by yangyebo 2017-12-12 上午11:21
 */
public abstract class HttpPostGeneratorFactory{

    public static HttpPostGenerator chooseGenerator(HttpRequest httpRequest) {

        if(httpRequest.getContentType().equals(HttpRequest.ContentType.JSON)) {
            return new JsonHttpPostGenerator();
        }

        return new FormHttpPostGenerator();
    }
}
