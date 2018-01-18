package cn.hydralisk.common.file;

import java.io.Serializable;

/**
 * 获取FTP文件基本信息
 * @author master.yang
 */
public class PullFileFtpInfo implements Serializable {

    private static final long serialVersionUID = -7180297758879562514L;

    private String ip; //ip

    private int port; //端口

    private String username; //用户名

    private String password; //密码

    private String path; //文件存放路径

    private String fileName; //文件名

    private String localStorageFile; //获取文件后存放本地路径

    //getter and setter
    public String getIp() {
        return ip;
    }

    public void setIp(String ip) {
        this.ip = ip;
    }

    public int getPort() {
        return port;
    }

    public void setPort(int port) {
        this.port = port;
    }

    public String getUsername() {
        return username;
    }

    public void setUsername(String username) {
        this.username = username;
    }

    public String getPassword() {
        return password;
    }

    public void setPassword(String password) {
        this.password = password;
    }

    public String getPath() {
        return path;
    }

    public void setPath(String path) {
        this.path = path;
    }

    public String getFileName() {
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public String getLocalStorageFile() {
        return localStorageFile;
    }

    public void setLocalStorageFile(String localStorageFile) {
        this.localStorageFile = localStorageFile;
    }
}
