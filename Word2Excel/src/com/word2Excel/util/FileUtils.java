package com.word2Excel.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;
import java.io.OutputStream;

import org.apache.poi.hwpf.extractor.WordExtractor;
/**
 * 
 * @author Administrator
 *
 */
public class FileUtils {
	/**
	 * 
	 * @param sourcePath
	 * @param targetPath
	 * @throws IOException
	 */
	public static void copyFile(String sourcePath, String targetPath) throws IOException{
		InputStream in = new FileInputStream(sourcePath);
		OutputStream out =  new FileOutputStream(targetPath);
		byte [] b = new byte[1024];
		int n = 0;
		while ((n= in.read(b))!=-1) {
			out.write(b,0,n);		
		}
		in.close();
		out.close();
		
	}
	/**
	 * 
	 * @param object
	 * @param filePath
	 * @throws IOException
	 * @throws ClassNotFoundException
	 */
	public static void writeObject(Object object,String filePath) throws IOException, ClassNotFoundException{
		File file = new File(filePath);  
        ObjectOutputStream oout = new ObjectOutputStream(new FileOutputStream(file));
        oout.writeObject(object);  
        oout.close(); 
	}
	/**
	 * 
	 * @param filePath
	 * @param clasz
	 * @throws FileNotFoundException
	 * @throws IOException
	 * @throws ClassNotFoundException
	 */
	public static void readObject(String filePath,Class clasz) throws FileNotFoundException, IOException, ClassNotFoundException{
		ObjectInputStream oin = new ObjectInputStream(new FileInputStream(filePath));  
        Object o= oin.readObject(); 
        oin.close();  
	}
	
	 public static String getTextFromWord(String filePath) {  
	        String result = null;  
	        File file = new File(filePath);  
	        FileInputStream fis = null;  
	        try {  
	            fis = new FileInputStream(file);   
	            @SuppressWarnings("resource")
				WordExtractor wordExtractor = new WordExtractor(fis); 
	            
	            result = wordExtractor.getText();  
	        } catch (FileNotFoundException e) {  
	            e.printStackTrace();  
	        } catch (IOException e) {  
	            e.printStackTrace();  
	        } finally {  
	            if (fis != null) {  
	                try {  
	                    fis.close();  
	                } catch (IOException e) {  
	                    e.printStackTrace();  
	                }  
	            }  
	        }  
	        return result;  
	    }  
}
