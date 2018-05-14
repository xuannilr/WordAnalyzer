package com.word2Excel.util;

import java.util.Collection;
import java.util.Map;

/**
 * 通用方法
 * @author Administrator
 *
 */
public class CommonUtils {
	
	public static boolean isNull(Object object){
		boolean flag = false;
		if(object == null){
			flag = true;
		}
		return flag;
	}
	
	public static boolean isNull(Collection<?> c){
		boolean flag = false;
		if(c == null||c.size()<=0){
			flag = true;
		}
		return flag;
	}
	public static boolean isNull(Map<?, ?> c){
		boolean flag = false;
		if(c == null||c.size()<=0){
			flag = true;
		}
		return flag;
	}
	
	public static boolean isNull(String str){
		boolean flag = false;
		if("".equals(str) || str == null){
			flag = true;
		}
		return flag;
	}
	public static int str2Int(String str){
		int i = 0 ;
		if(!isNull(str)&&isNum(str)){
			i = Integer.parseInt(str);
		}
		return i;
	}
	public static boolean isNum(String str){
		boolean flag = false;
		if(str.matches("^[0-9]*$")){
			flag = true;
		}
		return flag;
	}
	
}
