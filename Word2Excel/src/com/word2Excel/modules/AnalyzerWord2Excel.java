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
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableCell;
import org.apache.poi.hwpf.usermodel.TableIterator;
import org.apache.poi.hwpf.usermodel.TableRow;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import com.word2Excel.bean.Bid;
import com.word2Excel.bean.Project;
import com.word2Excel.util.CommonUtils;
import com.word2Excel.util.Constants;

public class AnalyzerWord2Excel {
	/**
	 * 
	 * @param target
	 * @param projects
	 */
	public void write2Excel(File target,List<Project> projects){
		InputStream in = null;
		Workbook workbook = null;
		try {
    		in =  new FileInputStream(target);
    		workbook =  WorkbookFactory.create(in);
    		Sheet sheet = null;
			sheet = (Sheet) workbook.getSheetAt(0);
    		int index = 1 ;
			for (Project p : projects) {				
    			Row row = sheet.getRow(index);
				if(row==null){
					row = sheet.createRow(index);
				}
				row.createCell(0).setCellValue( p.getpDate());
				row.createCell(1).setCellValue( p.getpSequence());
				row.createCell(2).setCellValue( p.getpCorporation());
				row.createCell(3).setCellValue( p.getpName());
				List <Bid> bids = p.getBids();
				if(!CommonUtils.isNull(bids)){
					int length = bids.size();
					
					for (int j = 0 ; j<length; j++) {
						index ++;
						Row r = sheet.getRow(index);
						if(r==null){
							r = sheet.createRow(index);
						}
						Bid bid = bids.get(j);
						if(bid!=null){							
							r.createCell(4).setCellValue(bid.getBidName());
							r.createCell(5).setCellValue(bid.getPrices());
						}
					}
				}else{
					++index;
				}
    			
			}	
			FileOutputStream fo = new FileOutputStream(target); // 输出到文件
	        workbook.write(fo);
    		in.close();
    		fo.close();
    		System.out.println("写入sheet表：" + workbook.getSheetName(0) + " 完成");
		} catch (EncryptedDocumentException | InvalidFormatException | IOException e) {
			e.printStackTrace();
		}
	}
	public void readFromWord(File target, File source){
		
	}
	/**
	 *  获取第一层目录
	 * @param source
	 * @return
	 */
	public ArrayList<File> getFirstFolder( File source){
		ArrayList<File> folders = new ArrayList<File>();
		File[] files = null; 
		if(source.isDirectory()){
			files = source.listFiles();
			for (File file : files) {
				if(file.isDirectory()){
					folders.add(file);
				}
			}
		}
		return folders;
		
	}
	/**
	 * 获取第二层目录
	 * @param parents
	 * @return
	 */
	public Map <String, List<File>> getSecondFolder(ArrayList<File> parents){
		Map <String, List<File>> folders = new HashMap<String, List<File>>();
		if(!CommonUtils.isNull(parents)){
			for (File parent : parents) {
				ArrayList <File> childrens = new ArrayList<File>();
				folders.put(parent.getName(), childrens);
				File [] fs = parent.listFiles();
				for (File child : fs) {
					if(child.isDirectory()){
						childrens.add(child);
					}
				}
			}
		}
		return folders;
	}
	
	/**
	 * 获取     指定目录/指定文件
	 * @param file
	 * @return
	 */
	
	public List<File> getFileByFilter(File parentFolder,String filter){
		List <File> folders = new ArrayList<File>();  //
		File[] files = parentFolder.listFiles();
		for (File file : files) {
			if(file.getName().indexOf(filter)!=-1){
				folders.add( file) ;
			}
		}
		return folders;
	}
	
	public void getFile(File parentFolder,List<File> children,String filter){
		if(parentFolder != null){
			File[] files = parentFolder.listFiles();
			if(files!=null&&files.length>0){				
				for (File file : files) {
					if(file.getName().indexOf(filter)!=-1){
						children.add( file) ;
					}else{
						getFile(file, children, filter);
					}
				}
			}
		}
	}
	public void generatePorject(File file, File tar){

		ArrayList<File> parents = new ArrayList<File>();
		List<Project> ps = new ArrayList<Project>();
		parents = getFirstFolder(file);
		
		for (File file2 : parents) {
			System.out.println(file2.getAbsolutePath());
		}
		
		Map<String ,List<File>> map  = new HashMap<String, List<File>>();
		map = getSecondFolder(parents);
		
		if(!CommonUtils.isNull(map)){
			for (Entry<String, List<File>> entry: map.entrySet()) {

				String time = entry.getKey();  //获取招标 时间
				System.out.println(time);
				
				List<File> project =  entry.getValue();
				for (File file2 : project) {
					Project p = new Project();
					p.setpDate(time);
					p.setpName(file2.getName());
					

					System.out.println("|------> "+file2.getName());
					File invitationFolder = getInvitationFolder(file2);  //招标目录
					if(invitationFolder!= null){						
						File doc = getInvitationDoc(invitationFolder);
						if(doc != null ){							
							String name = doc.getName();
							String no  = analysisString(doc, "招标编号",Constants.PATTERN);
							String inva = analysisString(doc,"招 标 人",Constants.PATTERN);
							System.out.println("|------------> "+name);
							System.out.println("|-----------------> 招标编号: "+no);
							System.out.println("|-----------------> 招 标 人 : "+inva);
							
							p.setpSequence(no);
							p.setpCorporation(inva);
						}else{
							//System.out.println("|------------> "+"******");
						}
					}
					List<File> tends = getTendersMainFolder( getTenderFolder(file2));
					for (File tenderFolder : tends) {
						System.out.println(tenderFolder.getName());
						Bid bid = new Bid();
						File[] docs =  tenderFolder.listFiles();
						
						if(!CommonUtils.isNull(docs)){
							
							for (File file3 : docs) {
								if(file3.isFile()&&file3.getName().endsWith(Constants.FileType.doc.getName())){
									if("".equals(bid.getBidName())){										
										String name = analysisString(file3,"投标人",Constants.PATTERN1);
										System.err.println("<<name ---"+name+"--->>");
										if(name.indexOf("有限公司")!=-1)bid.setBidName(name);
									}
									if("".equals(bid.getPrices()) ){
										String prices  = analysisString(file3,"投标总价",Constants.PATTERN1);
										System.err.println("<< prices ---"+prices+"--->>");
										bid.setPrices(prices);;
									} 
//									bid.setBidName(analysisString(file3,"投标总价",Constants.PATTERN1));
//									bid.setBidName(analysisString(file3,"投标机型",Constants.PATTERN1));
//									bid.setBidName(analysisString(file3,"轮毂高度",Constants.PATTERN1));
//									bid.setBidName(analysisString(file3,"单价",Constants.PATTERN1));
								}
							}
						}
						p.getBids().add(bid);
					}
					ps.add(p);
				}
			}
		}
		
		write2Excel(tar, ps);
	
		
	} 
	
	/**
	 * 获得  招标目录
	 * @param parent
	 * @return
	 */
	public File getInvitationFolder(File parent){
		return getFileByFilter(parent, Constants.TYPE_INVITATION_FOR_BIDS).get(0);
	}
	
	/**
	 * 获得  投标目录
	 * @param parent
	 * @return
	 */
	public File getTenderFolder(File parent){
		return getFileByFilter(parent, Constants.TYPE_TENDER).get(0);
	}
	
	public List<File> getTendersMainFolder(File parent){
		List<File> children  =new ArrayList<File>();
		List<File> tenders = new ArrayList<File>();
		Map <String,File> map = new HashMap<String, File>();
		getFile(parent, children,Constants.FileType.doc.getName());
		if(!CommonUtils.isNull(children)){			
			for (File file : children) {
				map.put(file.getParent(),file.getParentFile());
			}
		}
		if(!CommonUtils.isNull(map)){			
			for (Entry<String, File> entry: map.entrySet()) {
				tenders.add(entry.getValue());
			}
		}
		return tenders;
	}
	
	
	
	
	/**
	 *  获得  招标文件
	 * @param parent
	 * @return
	 */
	public File getInvitationDoc(File parent){
		List <File> children = new ArrayList<File>();
		File doc =  null;
		getFile(parent,children , Constants.TYPE_BUSINESS);
		if(!CommonUtils.isNull(children)){
			for (File file : children) {
				if(file.getName().toLowerCase().endsWith(Constants.FileType.doc.getName())||
						file.getName().toLowerCase().endsWith(Constants.FileType.docx.getName())){
					doc = file;
					break;
				}
			}
		}
		return  doc;
	} 
	
	/**
	 * 获取文档的  text 文本
	 * @param doc
	 * @return
	 */
	public List<String> getAllTextFromWord( HWPFDocument doc) {
		/*if (path.endsWith(".doc")) {  
            InputStream is = new FileInputStream(new File(path));  
            WordExtractor ex = new WordExtractor(is);  
            buffer = ex.getText();  
            ex.close();  
        } else if (path.endsWith("docx")) {  
            OPCPackage opcPackage = POIXMLDocument.openPackage(path);  
            POIXMLTextExtractor extractor = new XWPFWordExtractor(opcPackage);  
            buffer = extractor.getText();  
            extractor.close();  
        } else {  
            System.out.println("此文件不是word文件！");  
        }  */
		
        String  content = doc.getDocumentText();
       
      
        String c[] = content.split("\r|\n");
        List<String> cs = new ArrayList<String>(); 
        for (String string : c) {
			if(string.indexOf(Constants.Splitor.colon_zh.getName())!=-1||string.indexOf(Constants.Splitor.colon.getName())!=-1){
				cs.add(string.trim());
			}
		}
		return cs;         
	}
	
	@SuppressWarnings("resource")
	public  String analysisString(File file,String keyword,String pattern){
		String c = "";
		try {
			
			InputStream in =  new FileInputStream(file);
			
			HWPFDocument doc = null;
			Range range  = null;
			List<String> strFromDocPara = null;
			XWPFDocument docx = null;
			if(file.isHidden()){
				return c;
			}
			if(file.getName().endsWith(Constants.FileType.doc.getName())){
				System.out.println("-****->"+file.getName());
				doc = new HWPFDocument(in);
				range = doc.getRange();
				strFromDocPara = getAllTextFromWord(doc);
			}else{
				return c;
			}
			Map<Integer, List<String>> strFromDocTable = getTableContentFromWord(range);	
			for (String string : strFromDocPara) {
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
			if("".equals(c)&&!CommonUtils.isNull(strFromDocTable)){
				Iterator<Entry<Integer,List<String> >> it = strFromDocTable.entrySet().iterator();
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
		} catch (IOException e) {
			e.printStackTrace();
		}catch (StringIndexOutOfBoundsException e) {
			e.printStackTrace();
		}
		
		return c;
		
	}
	/**
     * 读文档中的表格
     * 
     * @param pTable
     * @param cr
     * @throws Exception
     */
    public void readTable(TableIterator it, Range rangetbl) throws Exception{
    	
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
			                map.put( k ,row.getCell(k).toString());
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
    
    public void writeData2Excel(){
    	
    	
    }
    /**
     * 从 word中解析 table 
     * @param rangetbl
     * @return
     */
    public Map<Integer,List<String >> getTableContentFromWord(Range rangetbl){
    	Map <Integer,List<String>> tabmap = new HashMap<Integer, List<String>>();
    	if(rangetbl ==null) return null;
    	try {   
            TableIterator it = new TableIterator(rangetbl);  
            int  index  = 0;
            while(it.hasNext()){  
            	List<String > tbContent = new ArrayList<String>();  
                Table tb = (Table)it.next();  
                for(int i = 0;i < tb.numRows();i++){                //获取  row
 
                    TableRow tr = tb.getRow(i);
                    StringBuilder sb = new StringBuilder("(");
                   
                    for(int j = 0;j < tr.numCells();j++){            //  获取 cell
                        TableCell td = tr.getCell(j);  
                        StringBuilder tdCon = new StringBuilder();
                        for(int k = 0;k < td.numParagraphs();k++){   //  获取 cell content
                            Paragraph para = td.getParagraph(k);
                            tdCon.append(para.text().trim());    
                        }
                        if(tdCon.length()>1){
                        	
                        	sb.append(tdCon.toString().replaceAll("：", "")); //替换中文字符":"
                        	sb.append( (j < tr.numCells() -1)? ":" : "");
                        }
                    }  
                    sb.append(")");
                    if(sb.length()>2){                    	
                    	tbContent.add(sb.toString());
                    }
                } 
                ++index;
                tabmap.put(index, tbContent);
            }  

        } catch (Exception e) {  
            e.printStackTrace();  
        } 
    	return tabmap;
    	
    }
	
}
