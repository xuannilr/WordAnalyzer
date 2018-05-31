package com.word2Excel.util;

import org.apache.log4j.Logger;

/**
 * 日志管理
 * 
 * @author cheng_shiming
 * 
 */
public class LoggerUtil {

	private Logger logger = null;

	/**
	 * 实例化日志管理
	 * 
	 * @param clazz
	 */
	@SuppressWarnings("rawtypes")
	public LoggerUtil(Class clazz) {
		logger = Logger.getLogger(clazz);
	}

	/**
	 * 记录日志
	 * 
	 * @param message
	 */
	public void log(Throwable t) {
		log("", t);
	}
	/**
	 * 记录日志
	 * 
	 * @param message
	 */
	public void log(String message) {
		logger.info(message);
		logger.error(message);
		logger.debug(message);
		logger.warn(message);
	}

	/**
	 * 记录日志
	 * 
	 * @param message
	 * @param t
	 */
	public void log(Object message, Throwable t) {
		logger.info(message, t);
		logger.error(message, t);
		logger.debug(message, t);
		logger.warn(message, t);
	}
}
