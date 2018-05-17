package com.word2Excel.modules;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.POIXMLDocument;
import org.apache.poi.POIXMLTextExtractor;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.HWPFOldDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableCell;
import org.apache.poi.hwpf.usermodel.TableIterator;
import org.apache.poi.hwpf.usermodel.TableRow;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlException;

import com.word2Excel.bean.CustomFile;
import com.word2Excel.util.CommonUtils;
import com.word2Excel.util.Constants;

public class POIUtils {
	/**
     * 从Excel中获取 数据  
     * @param excel
     * @return
     * @throws FileNotFoundException
     * @throws IOException
     * @throws EncryptedDocumentException
     * @throws InvalidFormatException
     */
    public static Map<Integer ,String> readDataFromExcel(File excel) {
    	Map<Integer ,String> map =new HashMap<Integer, String>();
    	try {
			InputStream in =  new FileInputStream(excel);
			Workbook workbook =  WorkbookFactory.create(in);
			Sheet sheet = null;
			 for (int i = 0; i < workbook.getNumberOfSheets(); i++) {// 获取每个Sheet表
			     sheet = (Sheet) workbook.getSheetAt(i);
			   
			     Row row = sheet.getRow(0);
			     if (row != null) {
			         for (int k = 0; k < row.getLastCellNum(); k++) {// getLastCellNum，是获取最后一个不为空的列是第几个
			             if (row.getCell(k) != null) { // getCell 获取单元格数据
			                map.put( k ,row.getCell(k).toString().replaceAll("\n",""));
			             } else {
			                 System.out.print("\t");
			             }
			         }
			     }
			     System.out.println("读取sheet表：" + workbook.getSheetName(i) + " 完成");
			 }
			 in.close();
		} catch (EncryptedDocumentException | InvalidFormatException | IOException e) {
			e.printStackTrace();
		}
    	return map;
    }
    
    
    /**
     * 从Excel中获取 数据  
     * @param excel
     * @return
     * @throws FileNotFoundException
     * @throws IOException
     * @throws EncryptedDocumentException
     * @throws InvalidFormatException
     */
    public static boolean writeData2Excel(File excel,Map<String, Map<Integer,List<String>>> data) {
		InputStream in = null;
		Workbook workbook = null;
		Boolean isSuccessful = false;
		try {
    		in =  new FileInputStream(excel);
    		workbook =  WorkbookFactory.create(in);
    		Sheet sheet = null;
			sheet = (Sheet) workbook.getSheetAt(0);
    		int rowIndex = 1 ;
    		Iterator<Entry<String, Map<Integer, List<String>>>> it = data.entrySet().iterator();
			while (it.hasNext()) {
				Entry<String, Map<Integer, List<String>>> entry = it.next();
				Map<Integer, List<String>> rowData = entry.getValue();
				Iterator<Entry<Integer, List<String>>> rowDataIt = rowData.entrySet().iterator();
				
				int maxRowindex  = 1;
				while (rowDataIt.hasNext()) {
					Entry<Integer, List<String>> en = rowDataIt.next();
					int colIndex = en.getKey();
					List<String> colData = en.getValue();
					int tempRowIndex = rowIndex;
					for ( String cellValue : colData) {
						Row row = sheet.getRow(tempRowIndex);
						if(row==null){
							row = sheet.createRow(tempRowIndex);
						}
						if(colIndex == -1){continue;}
						Cell cell= row.createCell(colIndex);
						cell.setCellValue(cellValue);
						tempRowIndex++;
						if(maxRowindex < tempRowIndex){
							maxRowindex = tempRowIndex;
						}
						
					}
				}
				rowIndex = maxRowindex;
				
			}
			FileOutputStream fo = new FileOutputStream(excel); // 输出到文件
	        workbook.write(fo);
    		in.close();
    		fo.close();
    		isSuccessful = true;
    		if(isSuccessful){
    			System.out.println("写入sheet表：" + workbook.getSheetName(0) + " 完成");
    		}else{
    			System.out.println("写入sheet表：" + workbook.getSheetName(0) + " 失败");
    		}
		} catch (EncryptedDocumentException | InvalidFormatException | IOException e) {
			e.printStackTrace();
		}
		return isSuccessful;
	}
    /**
     * 分析纯文本内容
     * @param strContainer
     * @param keyword
     * @param pattern
     * @return
     */
    public static String analysisString(List<String> strContainer, String keyword, String pattern){
    	String c = "";
    	
    	for (String string : strContainer) {
			if(string.indexOf(keyword)!=-1&&string.trim().matches(pattern)){
				String temp = string.trim();
				if(temp.matches(Constants.PATTERN)){
					temp = temp.substring(1, string.length()-1);
				}
				String s[] = 
						temp.indexOf(Constants.Splitor.colon_zh.getName())==-1?
								temp.split(Constants.Splitor.colon.getName()):
									temp.split(Constants.Splitor.colon_zh.getName());
				if(!CommonUtils.isNull(s)&&s.length>0){
					
					c = s[1];
					if(c.length()>30){
						c = "";
						continue;
					}
					break;
				}
			}else{
				continue;
			}
		}
    	
    	return c;
    }
    /**
     * 分析 表格中的内容
     * @param strContainer
     * @param keyword
     * @param pattern
     * @return
     */
    public static String analysisString(Map<Integer, List<String>> strContainer, String keyword, String pattern){
    	String c = "";
    	if(!CommonUtils.isNull(strContainer)){
			Iterator<Entry<Integer,List<String> >> it = strContainer.entrySet().iterator();
			while (it.hasNext()) {
				List<String> trows = it.next().getValue();
				for (String string : trows) {
					if(string.indexOf(keyword)!=-1&&string.trim().matches(pattern)){
						String temp = string.substring(1, string.length()-1);
						String s[] = temp.split(Constants.Splitor.colon.getName());
						if(!CommonUtils.isNull(s)&&s.length>1){
							c = s[1];
							break;
						}
					}else{
						continue;
					}
				}
			}
		}
    	return c;
    }

   
	  /**
		 * 获取文档的  text 文本
		 * @param doc
		 * @return
		 */
		public static List<String> getAllTextFromWord(String path) {
			List<String> cs = new ArrayList<String>();
			try {
				String buffer = "";
				
				InputStream in = new FileInputStream(new File(path));  
				if (path.endsWith(".doc")) {  
				    WordExtractor ex = new WordExtractor(in); 
				    buffer = ex.getText();  
				    ex.close();  
				} else if (path.endsWith("docx")) {  
				   // OPCPackage opcPackage = POIXMLDocument.openPackage(path);   //容易出问题   Zip bomb detected!
				    XWPFDocument document = new XWPFDocument(in);
				    POIXMLTextExtractor extractor = new XWPFWordExtractor(document); 
				    buffer = extractor.getText();  
				    extractor.close();  
				} else {  
				    System.err.println("此 [ "+ path +" ] 不是word文件！");  
				}  
				String c[] = buffer.split("\r|\n");
				
				for (String string : c) {
//					if(string.indexOf(Constants.Splitor.colon_zh.getName())!=-1
//							||string.indexOf(Constants.Splitor.colon.getName())!=-1){
//					}
					if("".equals(string)||string.length()>30){
						continue;
					}
					cs.add(string.trim().replaceAll("\\s", ""));
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
			return cs;         
		}
		
		/**
		 *   重写word 操作
		 *   
		 */
		public Map<Integer,Map<Integer,String>> getTextExtractor(String path){
			String buffer = "";
			try {
				InputStream in = new FileInputStream(new File(path));  
				if (path.endsWith(".doc")) {  
					HWPFDocument word2003 =new  HWPFDocument(in);
					WordExtractor ex = new WordExtractor(word2003);
				    buffer = ex.getText();  
				    ex.close();  
				} else if (path.endsWith("docx")) {  
				    XWPFDocument word2007 = new XWPFDocument(in);
				    POIXMLTextExtractor extractor = new XWPFWordExtractor(word2007); 
				    buffer = extractor.getText();  
				    extractor.close();  
				} else {  
				    System.err.println("此 [ "+ path +" ] 不是word文件！");  
				}
			} catch (IOException e) {
				e.printStackTrace();
			}  
			
			return null;
		}
		
		/**
		 * convert word2003 table  to list
		 */
		
		public List<Map<String,String>> convert2003Table(HWPFDocument word2003){
			Range range = word2003.getRange();
	    	List<Map<String,String>> tabList = new ArrayList<Map<String,String>>();
	    	if(range ==null) return null;
	    	try {   
	            TableIterator it = new TableIterator(range);  
	            while(it.hasNext()){  
	            	Map<String,String > tbContent = new HashMap<String, String>();  
	                Table tb = (Table)it.next();  
	                for(int i = 0;i < tb.numRows();i++){                //获取  row
	 
	                    TableRow tr = tb.getRow(i);
	                   
	                    for(int j = 0;j < tr.numCells();j++){            //  获取 cell
	                        TableCell td = tr.getCell(j);  
	                        StringBuilder tdCon = new StringBuilder();
	                        for(int k = 0;k < td.numParagraphs();k++){   //  获取 cell content
	                            Paragraph para = td.getParagraph(k);
	                            tdCon.append(para.text().trim());    
	                        }
	                        String value = tdCon.toString().replaceAll("：", "").replaceAll("\n", ""); //替换中文字符":" , 换行"\n"
	                        String key = i+","+j;       //  行列坐标
	                        tbContent.put(key, value); 
	                    }  
	                } 

	                tabList.add(tbContent);
	            }  

	        } catch (Exception e) {  
	            e.printStackTrace();  
	        } 
	    	
	    	
	    
			return tabList;
		}
		
		/**
		 * convert word2007 table  to list
		 */
		public List<Map<String,String>> convert2007Table(XWPFDocument word2007){
			List <XWPFTable> tables = word2007.getTables();
			List <Map<String,String>> tableList  = new ArrayList<Map<String,String>>(); //ready to return
			for (XWPFTable xwpfTable : tables) {
				Map <String, String > tabMap = new HashMap<String, String>();
				int maxRow = xwpfTable.getRowBandSize();
				for(int row = 0 ; row<maxRow ;row++){
					XWPFTableRow tablerow = xwpfTable.getRow(row);
					int  maxCol = tablerow.getTableCells().size();
					for(int col = 0 ;col < maxCol;col++){
						XWPFTableCell cell = tablerow.getCell(col);
						List <XWPFParagraph> paragragh = cell.getParagraphs();
						StringBuilder sb = new StringBuilder();
						String value =sb.toString();
						String key = row+","+col;
						for (XWPFParagraph xwpfParagraph : paragragh) {
							sb.append(xwpfParagraph.getParagraphText().replaceAll("\\s", ""));
						}
						tabMap.put(key, value);
					}
				}
				tableList.add(tabMap);
			}
			return tableList;
		}
}
