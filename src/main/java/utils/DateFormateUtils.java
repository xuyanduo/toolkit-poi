package utils;

import java.text.ParsePosition;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * @Auther: xyd
 * @Date: 2019/7/11 11:21
 * @Description:
 */

public class DateFormateUtils {

    /**
     * @Description  日期转字符串
     * @Param [date, pattern] 日期  转换格式
     * @return  字符串
     **/
    public static String date2String(Date date,String pattern){
        if(date == null){
            return "";
        }
        SimpleDateFormat format = new SimpleDateFormat(pattern);
        return format.format(date);
    }

    /**
     * @Description  时间戳转换为字符串
     * @Param [seconds, format] 时间戳单位为秒 转换格式
     * @return 转换后的时间字符串
     **/
    public static String timeStamp2Date(String seconds,String format) {
        if(seconds == null || seconds.isEmpty() || seconds.equals("null")){
            return "";
        }
        if(format == null || format.isEmpty()) format = "yyyy-MM-dd HH:mm:ss";
        SimpleDateFormat sdf = new SimpleDateFormat(format);
        return sdf.format(new Date(Long.valueOf(seconds+"000")));
    }

    /**
     * @Description  毫秒时间戳转换为字符串
     * @Param [milliSecond, format] 单位为毫秒 转换格式
     * @return 转换后的时间字符串
     **/
    public static String milliSecond2Date(String milliSecond, String format) {
        if (milliSecond == null || milliSecond.isEmpty() || milliSecond.equals("null")) {
            return "";
        }
        if (format == null || format.isEmpty()) format = "yyyy-MM-dd HH:mm:ss";
        SimpleDateFormat sdf = new SimpleDateFormat(format);
        return sdf.format(new Date(Long.valueOf(milliSecond)));
    }


    /**
     * @Description  String 转换为 Date
     * @Param [strDate, pattern] 字符串 正则
     * @return  时间
     **/
    public static Date strToDate(String strDate, String pattern) {
        SimpleDateFormat formatter = new SimpleDateFormat(pattern);
        ParsePosition pos = new ParsePosition(0);
        Date strtodate = formatter.parse(strDate, pos);
        return strtodate;
    }
}
