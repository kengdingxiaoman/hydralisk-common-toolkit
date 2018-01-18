package cn.hydralisk.common.file;

import org.apache.commons.net.ftp.FTPFile;
import org.apache.commons.net.ftp.FTPFileFilter;

/**
 * 指定FTP文件名的filter
 * @author master.yang
 */
public class SpecifiedFTPFileFilter implements FTPFileFilter {

    private String specifiedFileName;

    public SpecifiedFTPFileFilter() {
        super();
    }

    public SpecifiedFTPFileFilter(String specifiedFileName) {
        this.specifiedFileName = specifiedFileName;
    }

    /** 
     * @see FTPFileFilter#accept(FTPFile)
     */
    @Override
    public boolean accept(FTPFile file) {
        return file.getName().equals(specifiedFileName);
    }
}
