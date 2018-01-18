package cn.hydralisk.common.file;

import cn.hydralisk.common.utils.CommonUtils;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.net.ftp.FTP;
import org.apache.commons.net.ftp.FTPClient;
import org.apache.commons.net.ftp.FTPFile;
import org.apache.commons.net.ftp.FTPReply;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;

/**
 * 通过FTP拉取文件
 * @author master.yang
 */
public class PullFileByFtpUtils {

    private static final Logger logger = LoggerFactory.getLogger(PullFileByFtpUtils.class);

    private final static int DATA_TIME_OUT_MILLISECOND = 1000 * 10;

    private final static int DEFAULT_BUFFER_SIZE = 1024;

    private final static String DEFAULT_ENCODING = "UTF-8";

    public File pullFile(PullFileFtpInfo pullFileFtpInfo) {
        logger.info("请求FTP服务器信息：{}", CommonUtils.showDetails(pullFileFtpInfo));

        FTPClient ftpClient = null;
        try {
            if (!paramsIsValid(pullFileFtpInfo)) {
                logger.info("FTP信息参数不合法，不请求FTP服务器.");
                return null;
            }

            return pull(pullFileFtpInfo, ftpClient);
        } catch (IOException ioEx) {
            logger.error("获取FTP服务器文件发生IO异常.", ioEx);
            return null;
        } catch (Exception ex) {
            logger.error("获取FTP服务器文件发生系统异常.", ex);
            return null;
        } finally {
            if (ftpClient != null && ftpClient.isConnected()) {
                try {
                    ftpClient.disconnect();
                } catch (IOException ioEx) {
                    logger.error("断开FTP服务器发生IO异常.", ioEx);
                }
            }
        }
    }

    private boolean paramsIsValid(PullFileFtpInfo pullFileFtpInfo) {
        if (pullFileFtpInfo == null) {
            return false;
        }

        if (StringUtils.isBlank(pullFileFtpInfo.getIp())
            || StringUtils.isBlank(pullFileFtpInfo.getUsername())
            || StringUtils.isBlank(pullFileFtpInfo.getPassword())
            || StringUtils.isBlank(pullFileFtpInfo.getFileName())
            || StringUtils.isBlank(pullFileFtpInfo.getLocalStorageFile())) {
            return false;
        }

        if (pullFileFtpInfo.getPort() <= 0) {
            return false;
        }

        return true;
    }

    private File pull(PullFileFtpInfo pullFileFtpInfo, FTPClient ftpClient) throws IOException,
                                                                           Exception {
        ftpClient = new FTPClient();
        ftpClient.connect(pullFileFtpInfo.getIp(), pullFileFtpInfo.getPort());
        ftpClient.setDataTimeout(DATA_TIME_OUT_MILLISECOND);

        int reply = ftpClient.getReplyCode();
        if (!FTPReply.isPositiveCompletion(reply)) {
            logger.info("FTP服务器没有应答：{}, 退出.", reply);
            return null;
        }

        boolean loginSuccess = ftpClient.login(pullFileFtpInfo.getUsername(),
            pullFileFtpInfo.getPassword());
        if (!loginSuccess) {
            logger.info("登陆FTP服务器失败");
            return null;
        }

        ftpClient.setBufferSize(DEFAULT_BUFFER_SIZE);//设置上传缓存大小
        ftpClient.setControlEncoding(DEFAULT_ENCODING);//设置编码
        ftpClient.setFileType(FTP.BINARY_FILE_TYPE);

        if (StringUtils.isNotBlank(pullFileFtpInfo.getPath())) {
            ftpClient.changeWorkingDirectory(pullFileFtpInfo.getPath());
        }

        ftpClient.enterLocalPassiveMode();
        FTPFile[] ftpFiles = ftpClient.listFiles(pullFileFtpInfo.getPath(),
            new SpecifiedFTPFileFilter(pullFileFtpInfo.getFileName()));
        if (ArrayUtils.isEmpty(ftpFiles)) {
            ftpClient.logout();
            logger.info("FTP服务器没有指定文件, 退出.");
            return null;
        }

        File outFile = new File(pullFileFtpInfo.getLocalStorageFile());
        File outFileParent = outFile.getParentFile();
        if (outFileParent != null && !outFileParent.exists()) {
            boolean mkdirSuccess = outFileParent.mkdirs();
            if (!mkdirSuccess) {
                logger.info("本地文件目录创建失败，退出.");
                return null;
            }
        }

        if (!outFile.exists()) {
            boolean createFileSuccess = outFile.createNewFile();
            if (!createFileSuccess) {
                logger.info("本地文件创建失败，退出.");
                return null;
            }
        }

        OutputStream outputStream = null;
        try {
            outputStream = new BufferedOutputStream(new FileOutputStream(outFile));
            ftpClient.retrieveFile(pullFileFtpInfo.getFileName(), outputStream);
        } catch (IOException ioEx) {
            logger.error("获取FTP文件发生IO异常.", ioEx);
            throw new IOException("get ftp file occur io excpetion!");
        } catch (Exception ex) {
            logger.error("获取FTP文件发生系统异常.", ex);
            throw new Exception("get ftp file occur system excpetion!");
        } finally {
            if (outputStream != null) {
                outputStream.close();
            }
        }

        ftpClient.logout();
        return outFile;
    }
}
