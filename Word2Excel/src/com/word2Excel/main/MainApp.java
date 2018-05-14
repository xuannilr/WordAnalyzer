package com.word2Excel.main;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Properties;



import com.word2Excel.bean.Bid;
import com.word2Excel.bean.Project;
import com.word2Excel.modules.AnalyzerWord2Excel;
import com.word2Excel.util.CommonUtils;
import com.word2Excel.util.Constants;

public class MainApp {

	public static void main(String[] args) throws FileNotFoundException, IOException {
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
		String path = (String) properties.get("sourceFilePath");
		String targetPath = (String) properties.get("targetFile");
		File target = new File("F://2015//new.xlsx");
		File file = new File(path);
		createPorject(file,target);	
		
		long end = System.currentTimeMillis();
		System.out.println();
		System.out.println("运行时间: "+ (end - start ));
	}
	
	public static void createPorject(File file,File tar) throws FileNotFoundException{
		AnalyzerWord2Excel eng = new AnalyzerWord2Excel();
		eng.generatePorject(file, tar);
	}

}
