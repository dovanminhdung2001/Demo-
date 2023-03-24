package org.example;

import org.example.util.DateUtils;

import java.text.ParseException;
import java.util.Date;

public class Main {
    public static void main(String[] args) throws ParseException {
        String date1 = "06-02-2023 10:55:45";
        String date2 = "06-02-2023 11:11:56";

        Date d1 = DateUtils.sdtf.parse(date1);
        Date d2 = DateUtils.sdtf.parse(date2);

        Long m1 = d1.getTime();
        Long m2 = d2.getTime();

        System.out.println(m1);
        System.out.println(m2);
        System.out.println(m2 - m1);

        Date d3 = new Date(1675699916604L );
        System.out.println(DateUtils.sdtf.format(d3));

        Date d4 = new Date(m2);
        System.out.println(DateUtils.sdtf.format(d4));

        System.out.println(m2 - 1675699916604L);
    }
}