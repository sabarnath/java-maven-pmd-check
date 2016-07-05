package com.saba;

import java.math.BigDecimal;
import java.text.DecimalFormat;

import org.apache.commons.lang.StringUtils;


public class DecimalCheck {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		System.out.println("priceConversion null :"+priceConversion(null));
    	System.out.println("priceConversion  :"+priceConversion(new BigDecimal(0.0)));
    	System.out.println("priceConversion  :"+priceConversion(new BigDecimal(989776559678.99)));
    	
    	System.out.println("priceCurrencyConversion null :"+priceCurrencyConversion(null));
    	System.out.println("priceCurrencyConversion  :"+priceCurrencyConversion(new BigDecimal(0.0)));
    	System.out.println("priceCurrencyConversion  :"+priceCurrencyConversion(new BigDecimal(9966885478.99)));


	}

    public static String priceConversion(BigDecimal price){
		
		  DecimalFormat formatter = new DecimalFormat("0.00");
		  return price != null ? formatter.format(price) : StringUtils.EMPTY;
	}
    
    public static String priceCurrencyConversion(BigDecimal bdPrice){

		  DecimalFormat formatter = new DecimalFormat("#,###,###,###,###,##0.00");
		  return bdPrice != null ? formatter.format(bdPrice) : StringUtils.EMPTY;
	}

}
