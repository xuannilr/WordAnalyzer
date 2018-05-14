package com.word2Excel.main;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;
import java.util.Properties;

import org.apache.poi.hwpf.HWPFDocument;

import com.word2Excel.bean.CustomFile;
import com.word2Excel.modules.AnalyzerWord2Excel;
import com.word2Excel.modules.FileAnalyzer;
import com.word2Excel.modules.POIUtils;

public class Main {
	
	public static void main(String[] args) {
		//File file = null;
//		file = new File("F:\\2015\\test.doc");
		//file = new File("F:\\2015");
//		try {
//			InputStream in = new FileInputStream(file);
//			HWPFDocument doc = new HWPFDocument(in);
//			AnalyzerWord2Excel analyzer = new AnalyzerWord2Excel();
//			List<String> all= analyzer.getAllTextFromWord(doc);
//			int i =0;
//			for (String string : all) {
//				System.out.println(string);
//				++i;
//				if (i== 10){
//					break;
//				}
//			}
//		} catch (FileNotFoundException e) {
//			e.printStackTrace();
//		} catch (IOException e) {
//			e.printStackTrace();
//		}
		
//		FileAnalyzer fileAnalyzer =  new FileAnalyzer();
//		List<CustomFile>list = fileAnalyzer.listAllCustomFile(file);
//		for (CustomFile customFile : list) {
//			System.out.println(printf("|"+"-----", customFile.getLevel())+"> "+customFile.getName());
//		}
		
long start = System.currentTimeMillis();
		
		Properties properties = new Properties();
		properties = new Properties();
		System.out.println(new MainApp().getClass());
		InputStream inputStream = new MainApp().getClass().getClassLoader().getResourceAsStream("config.properties");
		try {
			properties.load(inputStream);
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		String xmlFile = (String) properties.get("roleMapping");
		String path = (String) properties.get("sourceFilePath");
		String targetPath = (String) properties.get("targetFile");
		File target = new File(targetPath);
		File file = new File(path);
		FileAnalyzer fa =  new FileAnalyzer(path, targetPath);
		Map<String, Map<Integer, List<String>>> map = fa.handleResult(fa.resolvingXml(xmlFile));
		
		System.out.println(map.toString());
		
		POIUtils.writeData2Excel(new File(targetPath), map);
		long end = System.currentTimeMillis();
		System.out.println();
		System.out.println("运行时间: "+ (end - start ));
		
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
