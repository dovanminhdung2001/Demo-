package excel1;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.temporal.ChronoUnit;
import java.util.Date;

public class DateUtils {
    public static SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
    public static SimpleDateFormat stf = new SimpleDateFormat("HH:mm:ss");
    public static SimpleDateFormat sdtf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
    public static SimpleDateFormat dateUpFile = new SimpleDateFormat("dd_MM_yyyy_hh_mm_ss");

    public static Date now() {
        return new Date();
    }

    public static Date today() throws ParseException {
        return  sdf.parse(sdf.format(new Date()));
    }

    public static Date tomorrow() throws ParseException {
        return sdf.parse(sdf.format(Date.from(new Date().toInstant().minus(-1, ChronoUnit.DAYS))));
    }

    public static String nowStr() {
        return sdtf.format(new Date());
    }

    public static String todayStr() {
        return sdf.format(new Date());
    }

    public static String tomorrowStr () throws ParseException {
        return sdf.format(Date.from(new Date().toInstant().minus(-1, ChronoUnit.DAYS)));
    }

    public static String dateUpFile() {
        return dateUpFile.format(new Date());
    }
}
