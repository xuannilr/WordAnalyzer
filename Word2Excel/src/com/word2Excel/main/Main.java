package com.word2Excel.main;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;
import java.util.Properties;

import org.apache.log4j.Logger;

import com.word2Excel.modules.FileAnalyzer;
import com.word2Excel.modules.POIUtils;

public class Main {
	
	public static void main(String[] args) {

		Logger logger = Logger.getLogger(Main.class);
		
		long start = System.currentTimeMillis();
		Properties properties = new Properties();
		properties = new Properties();
		System.out.println(new Main().getClass());
		InputStream inputStream = new Main().getClass().getClassLoader().getResourceAsStream("config.properties");
		try {
			properties.load(inputStream);
		} catch (IOException e) {
			e.printStackTrace();
		}
		String xmlFile = (String) properties.get("roleMapping");
		String path = (String) properties.get("sourceFilePath");
		String targetPath = (String) properties.get("targetFile");
		FileAnalyzer fa =  new FileAnalyzer(path, targetPath);
		Map<String, Map<Integer, List<String>>> map = fa.handleResult(fa.resolvingXml(xmlFile));
		
		POIUtils.writeData2Excel(new File(targetPath), map);
		long end = System.currentTimeMillis();
		System.out.println();
		System.out.println("运行时间: "+ (end - start ));
		
		System.out.println(map.toString());
		logger.info(map.toString());
	}
	
	public static String printf(String ch,int num){
		StringBuilder sb = new StringBuilder();
		if(num>0){
			for (int i = 0; i < num; i++) {
				sb.append(ch);
			}
		}
		return sb.toString();
	}
}
