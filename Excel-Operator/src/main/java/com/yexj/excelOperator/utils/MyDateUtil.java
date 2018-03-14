package com.yexj.excelOperator.utils;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Description:
 *
 * @author ye.xinjian
 * @version 1.0.0
 * @date: 2017-01-10 17:29
 */
public final class MyDateUtil {
    private static Logger logger = LoggerFactory.getLogger(MyDateUtil.class);

    private static long ONE_MINUTE = 60 * 1000;
    private static long ONE_DAY = 24 * 60 * 60 * 1000;

    public static final String datePattern = "\\d{4}\\d{2}\\d{2}";
    public static final String buildingCodePattern = "[a-zA-Z]";

    public static final String YMD = "yyyyMMdd";
    public static final String Y_M = "yyyy-MM";
    public static final String Y_M_D = "yyyy-MM-dd";
    public static final String Y_M_DHHmmss = "yyyy-MM-dd HH:mm:ss";
    public static final String YYYY = "yyyy";

    public static final String Y_M_D_HH_mm_ss = "yyyy_MM_dd_HH_mm_ss";

    public static final String formatYYYY(Date date) {
        if (date == null) return null;
        return new SimpleDateFormat(YYYY).format(date);
    }

    public static String formatYMD(Date date) {
        if (date == null) {
            return null;
        }
        return new SimpleDateFormat(YMD).format(date);
    }

    public static Date parseYMD(String ymd) throws ParseException {
        if (ymd == null) return null;
        return new SimpleDateFormat(YMD).parse(ymd);
    }

    public static String formatY_M(Date date) {
        if (date == null) {
            return null;
        }
        return new SimpleDateFormat(Y_M).format(date);
    }

    public static String formatY_M_D(Date date) {
        if (date == null) {
            return null;
        }
        return new SimpleDateFormat(Y_M_D).format(date);
    }

    public static String formatY_M_DHHMMSS(Date date) {
        if (date == null) {
            return null;
        }
        return new SimpleDateFormat(Y_M_DHHmmss).format(date);
    }

    public static String formatY_M_D_HH_mm_ss(Date date) {
        if(date == null) {
            return null;
        }
        return new SimpleDateFormat(Y_M_D_HH_mm_ss).format(date);
    }

    public static Date parseY_M_DHHMMSS(String date) throws ParseException {
        if (date == null) return null;
        return new SimpleDateFormat(Y_M_DHHmmss).parse(date);
    }

    public static Date parseY_M_D(String date) {
        if (date == null) {
            return null;
        }
        Date resultDate = null;
        try {
            resultDate = new SimpleDateFormat(Y_M_D).parse(date);
        } catch (ParseException e) {
            logger.error("MyDateUtil 解析日期异常:{}", e);
        }
        return resultDate;
    }


    public static int diffMinuteNow(Date createTime) {
        long time = new Date().getTime() - createTime.getTime();
        time = time / ONE_MINUTE;
        return (int) time;
    }

    /**
     * soureDate 比 targetDate 大几天 忽略时分秒
     *
     * @param soureDate
     * @param targetDate
     * @return
     */
    public static int diffDay(Date soureDate, Date targetDate) {
        SimpleDateFormat dft = new SimpleDateFormat("yyyy-MM-dd");
        long time = 0;
        try {
            Date startDate = dft.parse(dft.format(soureDate.getTime()));
            Date endDate = dft.parse(dft.format(targetDate.getTime()));
            time = startDate.getTime() - endDate.getTime();
        } catch (ParseException e) {
        }
        time = time / ONE_DAY;
        return (int) time;
    }

    public static Date getOldDate(Date one, Date two) {
        if (one == null && two != null) return two;
        if (one != null && two == null) return one;
        if (one == null && two == null) return null;
        long diff = one.getTime() - two.getTime();
        if (diff >= 0) return one;
        else return two;
    }

    public static Date getEarlyDate(Date one, Date two) {
        if (one == null && two != null) return two;
        if (one != null && two == null) return one;
        if (one == null && two == null) return null;
        long diff = one.getTime() - two.getTime();
        if (diff >= 0) return two;
        else return one;
    }

    public static void main(String args[]) throws ParseException {
        SimpleDateFormat dft = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
         /*Date beginDate = new Date();
        Calendar date = Calendar.getInstance();
        date.setTime(beginDate);
        Date startDate = dft.parse(dft.format(date.getTime()));
        date.set(Calendar.DATE, date.get(Calendar.DATE) -100);
        Date endDate = dft.parse(dft.format(date.getTime()));*/
//        Date date1 = dft.parse("2017-02-28 15:00:01");
//        System.err.println(MyDateUtil.compareLastDay(date1));
        System.out.println(subductionMonths(parseY_M_D("2017-06-01"),1));
    }

    public static String getYearByDate(Date date) {
        Calendar c = Calendar.getInstance();
        c.setTime(date);
        return String.valueOf(c.get(Calendar.YEAR));
    }

    /**
     * 获取过账时间
     * B端每月最后一天下午4点30后收款，凭证延后至第二天做进
     *
     * @param date
     * @return
     */
    public static boolean compareLastDay(Date date) {
        Calendar nowCalendar = Calendar.getInstance();
        nowCalendar.setTime(date);
        final int nowDay = nowCalendar.get(Calendar.DAY_OF_MONTH);
        final int lastDay = nowCalendar.getActualMaximum(Calendar.DAY_OF_MONTH);
        if (nowDay - lastDay == 0) {//最后一天
            Date lastDate = getLastDayOfMonth(date);
            if (date.after(lastDate)) {
                return true;
            }
        }
        return false;
    }

    /**
     * 每月最后一天15:00以后B端不允许支付
     *
     * @param sDate1
     * @return
     */
    public static Date getLastDayOfMonth(Date sDate1) {
        Calendar cDay1 = Calendar.getInstance();
        cDay1.setTime(sDate1);
        final int lastDay = cDay1.getActualMaximum(Calendar.DAY_OF_MONTH);
        Date lastDate = cDay1.getTime();
        lastDate.setDate(lastDay);
        lastDate.setHours(15);
        lastDate.setMinutes(0);
        lastDate.setSeconds(0);
        return lastDate;
    }


    /**
     * 获取当前时间的零点或者23:59：59
     * flag为0表示凌晨00:00:00 flag为1表示凌晨23:59:59
     *
     * @param date
     * @param flag
     * @return
     */
    public static Date weeHours(Date date, int flag) {
        Calendar cal = Calendar.getInstance();
        cal.setTime(date);
        int hour = cal.get(Calendar.HOUR_OF_DAY);
        int minute = cal.get(Calendar.MINUTE);
        int second = cal.get(Calendar.SECOND);
        //时分秒（毫秒数）
        long millisecond = hour * 60 * 60 * 1000 + minute * 60 * 1000 + second * 1000;
        //凌晨00:00:00
        cal.setTimeInMillis(cal.getTimeInMillis() - millisecond);

        if (flag == 0) {
            return cal.getTime();
        } else if (flag == 1) {
            //凌晨23:59:59
            cal.setTimeInMillis(cal.getTimeInMillis() + 23 * 60 * 60 * 1000 + 59 * 60 * 1000 + 59 * 1000);
        }
        return cal.getTime();
    }


    /**
     * 取某一段时间内所有的月份
     *
     * @param beginTime 开始时间
     * @param endTime   结束时间
     * @return
     */
    public static List<String> getMonthList(String beginTime, String endTime) {
        SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");
        SimpleDateFormat monthFormat = new SimpleDateFormat("yyyy-MM");
        List<String> monthList = new ArrayList<String>();
        try {
            Date begin = format.parse(beginTime);
            Date end = format.parse(endTime);
            int months = (end.getYear() - begin.getYear()) * 12
                    + (end.getMonth() - begin.getMonth());

            for (int i = 0; i <= months; i++) {
                Calendar calendar = Calendar.getInstance();
                calendar.setTime(begin);
                calendar.add(Calendar.MONTH, i);
                calendar.add(Calendar.DAY_OF_MONTH, 0);
                monthList.add(monthFormat.format(calendar.getTime()));
            }

        } catch (ParseException e) {
            e.printStackTrace();
        }

        return monthList;
    }

    /**
     * 取某一个月的第一天
     *
     * @param date
     * @return
     */
    public static Date getFirstDayByDate(String date) {
        SimpleDateFormat format = new SimpleDateFormat("yyyy-MM");
        Calendar calendar = Calendar.getInstance();
        try {
            Date begin = format.parse(date);
            calendar.setTime(begin);
            calendar.add(Calendar.DAY_OF_MONTH, 0);
        } catch (ParseException e) {
            e.printStackTrace();
        }
        return calendar.getTime();
    }

    /**
     * 计算延迟若干年后的时间戳
     *
     * @param years
     * @return
     */
    public static Date getTimestampWithDelayYears(int years) {
        Calendar calendar = Calendar.getInstance();
        calendar.add(Calendar.YEAR, years);
        calendar.set(Calendar.HOUR_OF_DAY, 23);
        calendar.set(Calendar.SECOND, 59);
        calendar.set(Calendar.MINUTE, 59);
        calendar.set(Calendar.MILLISECOND, 0);
        return new Date(calendar.getTime().getTime());
    }


    /**
     * 在某个时间上增加天
     *
     * @param date 时间
     * @param n    天
     * @return
     */
    public static Date addDay(Date date, int n) {
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(date);
        calendar.add(Calendar.DATE, n);//增加一天
        return new Date(calendar.getTime().getTime());
    }

    /**
     * 在某个时间上增加月
     *
     * @param date 时间
     * @param n    月
     * @return
     */
    public static Date addMonth(Date date, int n) {
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(date);
        calendar.add(Calendar.MONTH, n);//增加一个月
        return new Date(calendar.getTime().getTime());
    }

    /**
     * 某一段时间内减去多少天
     * @param date
     * @param n
     * @return
     */
    public static Date subductionDays(Date date,int n){
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(date);
        calendar.set(Calendar.DATE, calendar.get(calendar.DATE) - n);
        return new Date(calendar.getTime().getTime());
    }

    /**
     * 某一个时间减去多少小时
     * @param date
     * @param n
     * @return
     */
    public static Date subductionHours(Date date,int n){
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(date);
        calendar.set(Calendar.HOUR, calendar.get(calendar.HOUR) - n);
        return new Date(calendar.getTime().getTime());
    }

    /**
     * 某一段时间内减去多少月
     * @param date
     * @param n
     * @return
     */
    public static Date subductionMonths(Date date,int n){
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(date);
        calendar.set(Calendar.MONTH, calendar.get(calendar.MONTH) - n);
        return new Date(calendar.getTime().getTime());
    }

    /**
     * 日期格式正则匹配
     * @param date
     * @return
     */
    public static Boolean datePatternYYYYMMDD(String date){
        Pattern pattern = Pattern.compile(datePattern);
        Matcher match = pattern.matcher(date);
        return match.matches();
    }

    /**
     * 场馆号正则匹配
     * @param buildingCode
     * @return
     */
    public static Boolean validateBuildingCodePattern(String buildingCode){
        Pattern pattern = Pattern.compile(buildingCodePattern);
        Matcher match = pattern.matcher(buildingCode);
        return match.matches();
    }


}
