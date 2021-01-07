package com.glodon;

import org.springframework.util.StringUtils;

import java.io.PrintWriter;
import java.io.StringWriter;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * @author zhangjingfei
 * 12.30.2018
 */
public class CommonUtils {

    public static String toString(Object o) {
        if (o == null) {
            return "";
        } else {
            return o.toString();
        }
    }

    public static String toString(Date date, String format) {
        SimpleDateFormat formatter = new SimpleDateFormat(format);
        return formatter.format(date);
    }

    public static String formatSerialNumber(long number) {
        String serial = "000000" + number;
        int length = serial.length();
        return serial.substring(length - 6);
    }

    public static String parseIDFromDisplayName(String displayName) {
        if (StringUtils.isEmpty(displayName)) {
            return "0";
        }

        int beginIndex = displayName.lastIndexOf("(") + 1;
        return displayName.substring(beginIndex, displayName.length() - 1);
    }

    public static String parseNameFromDisplayName(String displayName) {
        if (StringUtils.isEmpty(displayName)) {
            return "空值";
        }

        int endIndex = displayName.lastIndexOf("(");
        if (endIndex < 0) {
            endIndex = displayName.length();
        }
        return displayName.substring(0, endIndex);
    }

    public static String exceptionToString(Throwable e) {
        StringWriter sw = new StringWriter();
        PrintWriter pw = new PrintWriter(sw, true);
        e.printStackTrace(pw);
        pw.flush();
        sw.flush();
        return sw.toString();
    }

    public static double formatDouble(double d, int scale) {
        //(1)
        // Math.round(d*100)/100;

        //(2)
        //NumberFormat nf = NumberFormat.getNumberInstance();
        //nf.setMaximumFractionDigits(2);
        //nf.setRoundingMode(RoundingMode.UP);
        //nf.format(d);

        //(3)
        //DecimalFormat df = new DecimalFormat("#.00");
        //df.format(d);

        //(4)
        //String.format("%.2f", d);

        //(5)
        BigDecimal bg = new BigDecimal(d).setScale(scale, RoundingMode.HALF_UP);
        return bg.doubleValue();
    }

    public static boolean isNumeric(String str) {
        for (int i = 0; i < str.length(); i++) {
            if (!Character.isDigit(str.charAt(i))) {
                return false;
            }
        }
        return true;
    }
}
