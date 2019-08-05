package utils;

import java.util.concurrent.ThreadLocalRandom;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class StringUtil {

	public static String getNullCharacterString() {
		return "";
	}

	public static String firstCharToLowerCase(String str) {
		return String.valueOf(str.charAt(0)).toLowerCase().concat(str.substring(1));
	}

	public static String firstCharToUpperCase(String str) {
		return String.valueOf(str.charAt(0)).toUpperCase().concat(str.substring(1));
	}

	public static boolean hasSpecifiedString(String str, String specified) {
		return str.indexOf(specified) >= 0;
	}

	/**
	 * 判断非空字符串是否是手机号格式
	 * 
	 * @param string
	 * @return
	 */
	public static boolean isTelephone(String string) {
		String regex = "^((13[0-9])|(14[5,7,9])|(15([0-3]|[5-9]))|(166)|(17[0,1,3,5,6,7,8])|(18[0-9])|(19[8|9]))\\d{8}$";
		if (string.length() != 11) {
			return false;
		}
		Pattern p = Pattern.compile(regex);
		Matcher m = p.matcher(string);
		boolean isMatch = m.matches();
		return isMatch;

	}

	/**
	 * 当字符串为空，或者仅有空格构成时返回true，否则返回false
	 * 
	 * @param value
	 * @return
	 */
	public static boolean isEmptyOrBlank(String value) {
		if (value == null) {
			return true;
		}
		value = value.trim();
		if (value.length() == 0) {
			return true;
		}
		return false;
	}

	public static boolean isNotEmptyOrBlank(String value) {
		if (isEmptyOrBlank(value)) {
			return false;
		}
		return true;
	}

	public static String randomNumber(Integer length) {
		StringBuilder builder = new StringBuilder(length);
		for (int i = 0; i < length; i++) {
			Integer random = ThreadLocalRandom.current().nextInt(0, 10);
			builder.append(random);
		}
		return builder.toString();
	}
}