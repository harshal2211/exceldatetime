package com.practice.exceldatetime;

import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.Date;
import java.util.TimeZone;

public class DateTimeExample {
	
	public static void main(String[] args) {
		System.out.println("Server Date Time now "+LocalDateTime.now(TimeZone.getTimeZone("GMT").toZoneId()));
		System.out.println("Local Date Time now "+LocalDateTime.now());
		System.out.println(new Date().toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime());
		
		
		Double d = 0.0041;
		System.out.println(d);
		
		Double div = 0.004d/100;
		System.out.println(div);
		
		System.out.println(Double.longBitsToDouble(Double.doubleToLongBits(div)));

		
		System.out.println(new BigDecimal(d).setScale(6, BigDecimal.ROUND_HALF_UP));
		System.out.println(d*60000);
	}

}
