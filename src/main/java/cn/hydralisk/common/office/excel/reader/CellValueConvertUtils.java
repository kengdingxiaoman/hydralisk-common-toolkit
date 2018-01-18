package cn.hydralisk.common.office.excel.reader;

import cn.hydralisk.common.utils.convert.DateUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.springframework.util.Assert;

import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * 转换单元格的工具类
 * created by master.yang 2018/1/18 下午3:09
 */
public abstract class CellValueConvertUtils {

    /**
     * 小数后两位是.00则忽略小数，否则保留小数
     *
     * 2.22222222222E11 -> 222222222222
     * 23.0 -> 23
     * 234.54 -> 234.54
     * 3.33333333389E9 -> 3333333333.89
     * 555555555 -> 555555555
     * 3.3333E-6 -> 0.0000033333
     * -888 -> -888
     *
     * @param numericCellValue
     * @return
     */
    public static String convertToNumberSmart(double numericCellValue) {
        BigDecimal bigDecimal = new BigDecimal(String.valueOf(numericCellValue));
        bigDecimal.setScale(2, BigDecimal.ROUND_HALF_UP);

        String bigDecimalStr = bigDecimal.toEngineeringString();

        NumberFormat nf = NumberFormat.getNumberInstance();
        nf.setMaximumFractionDigits(2);
        nf.setRoundingMode(RoundingMode.HALF_UP);
        nf.setGroupingUsed(false);
        String result = nf.format(bigDecimal);

        if (Double.valueOf(result) == 0) {
            //处理类似0.00034的情况
            return bigDecimalStr;
        } else {
            return result;
        }
    }

    public static String convertFormulaCellValue(Cell cell) {
        try {
            return convertToNumberSmart(cell.getNumericCellValue());
        } catch (IllegalStateException e) {
            return String.valueOf(cell.getStringCellValue());
        }
    }

    public static final String format(Date date) {
        Assert.notNull(date, "dateStr is null.");
        return new SimpleDateFormat(DateUtils.LONG_DATE_FORMAT).format(date);
    }
}
