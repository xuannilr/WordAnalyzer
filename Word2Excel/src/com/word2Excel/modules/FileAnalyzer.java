package com.word2Excel.modules;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;
import org.xml.sax.EntityResolver;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;

import com.word2Excel.bean.CustomFile;
import com.word2Excel.bean.vo.Enumeration;
import com.word2Excel.bean.vo.Group;
import com.word2Excel.bean.vo.ProjectItem;
import com.word2Excel.bean.vo.Thead;
import com.word2Excel.util.CommonUtils;
import com.word2Excel.util.Constants;
import com.word2Excel.util.LoggerUtil;

/**
 * 
 * @author li_ran
 *
 */
public class FileAnalyzer {
	private File root;
	private File excel;
	Map<Integer, String> dataFromExcel = null;
	LoggerUtil logger = null;
	private FileAnalyzer(){
		this.logger =  new LoggerUtil(this.getClass());
	}
	public FileAnalyzer(String path,String excelPath){
		this();
		this.root = new File(path);
		this.excel = new File(excelPath);
		this.dataFromExcel = POIUtils.readDataFromExcel(this.excel);
	}
	/**
	 * 将file 下 所有文件转化 为 自定义文件  
	 * @param file   
	 * @return
	 */
	public List <CustomFile> listAllCustomFile(File file){
		List <CustomFile> allFile = new ArrayList<CustomFile>();
		int level = 0;
		CustomFile customFile = new CustomFile();
		customFile.setLevel(level);
		customFile.setId(0);
		customFile.setName(file.getName());
		customFile.setAbsolutePath(file.getAbsolutePath());
		customFile.setFolder(file.isDirectory());
		allFile.add(customFile);
		level++;
		listAllFile(allFile, file ,level,customFile);
		return allFile;
	}
	/**
	 * 
	 * @param ruleMapping
	 */
	@SuppressWarnings("unchecked")
	public List<ProjectItem> resolvingXml(String ruleMapping){
		List< ProjectItem> pis = new ArrayList<ProjectItem>();
		try {
			SAXReader saxReader = new SAXReader();
			saxReader.setValidation(false);
			saxReader.setEntityResolver(new EntityResolver() {
				public InputSource resolveEntity(String publicId, String systemId)
						throws SAXException, IOException {
					return new InputSource(new ByteArrayInputStream(
							"<?xml version='1.0' encoding='utf-8'?>".getBytes()));
				}
			});
			saxReader.setEncoding("UTF-8");
			Document document = saxReader.read(this.getClass().getResourceAsStream("/" + ruleMapping));
			Element root = document.getRootElement();
			List<Element> projects = root.elements("project");
			if(!CommonUtils.isNull(projects)){
				
				for (Element element : projects) {
					ProjectItem pi = new ProjectItem();
					List<Element> group = element.elements("group");
					for (Element element2 : group) {
						pi.getGroups().add(new Group(element2));
					}
					pis.add(pi);
				}
			}
		} catch (DocumentException e) {
			e.printStackTrace();
		} 
		return pis;
	}
	
	/**
	 * 
	 * @param pis
	 * @throws IOException 
	 * @throws InvalidFormatException 
	 * @throws FileNotFoundException 
	 * @throws EncryptedDocumentException 
	 */
	
	public Map<String, Map<Integer, List<String>>> handleResult(List<ProjectItem> pis) {
		Map <String ,Map<Integer,List<String>>> map  = new HashMap<String, Map<Integer,List<String>>>();
		List<CustomFile> mainFolders =  getMainFolder(Constants.TYPE_INVITATION_FOR_BIDS);
		
		for (CustomFile customFile : mainFolders) {
			System.out.println(customFile.getName());
			CustomFile invatation = getCustomFileByName(customFile.getChildren(),Constants.TYPE_INVITATION_FOR_BIDS); //招标目录
			CustomFile tender = getCustomFileByName(customFile.getChildren(),Constants.TYPE_TENDER); //  投标目录
			Map<Integer, List<String>> ready2writing = new HashMap<Integer, List<String>>();	
			for (ProjectItem projectItem : pis) {
				List<Group> groups = projectItem.getGroups();
				for (Group group : groups) {
					if(Constants.TYPE_INVITATION_FOR_BIDS.equals(group.getKey())){  //01-招标文件
						List<CustomFile> docFiles =  getDocsByName(invatation);
						List<String> ready2AnalyParagraphs = new ArrayList<String>();
						List<Map<String,String>> ready2AnalyTables = new ArrayList<Map<String,String>>();
						for(CustomFile doc :docFiles){
							ready2AnalyParagraphs.addAll(doc.getParagrathsText()); 
							ready2AnalyTables.addAll(doc.getTablesParagraphsText());
						}
						List<Thead > tenderThs = group.getTheads();
						for (Thead thead : tenderThs) {
							analyzerRules(thead,ready2writing, ready2AnalyTables, ready2AnalyParagraphs,invatation);
						}
					}
					else{   //02-投标文件
						List<CustomFile> tenderFiles = getCustomFileByLevelOffset(tender,2); 
						List<Thead > tenderThs = group.getTheads();
						
						for(CustomFile tf :tenderFiles){  ////
							if(tf.isFolder()){
								List<CustomFile> docFiles =  getDocsByName(tf);
								System.out.println("file num-->"+docFiles.size());
								List<String> ready2AnalyParagraphs = new ArrayList<String>();
								List<Map<String,String>> ready2AnalyTables = new ArrayList<Map<String,String>>();
								for (CustomFile doc : docFiles) {  //取得所有 待解析字符集合
									ready2AnalyParagraphs.addAll(doc.getParagrathsText());
									ready2AnalyTables.addAll(doc.getTablesParagraphsText());
								}	
								for (Thead thead : tenderThs) {
									analyzerRules(thead, ready2writing, ready2AnalyTables, ready2AnalyParagraphs,null);
								}	
							}
						}
					}
				}
				
			}
			map.put(customFile.getName(),ready2writing );
		}
		return map;
	}
	private void analyzerRules(Thead thead,Map<Integer, List<String>> ready2writing,List<Map<String,String>> ready2AnalyTables,List<String>   ready2AnalyParagraphs,CustomFile invatation){
		int key = getMapKeyByValue(dataFromExcel, thead.getTitle());
		List<String> templist =  ready2writing.get(key);
		if(CommonUtils.isNull(templist)){
			templist = new ArrayList<String>();
			ready2writing.put(key, templist);
		}
		if(Constants.RuleType.folder.getName().equals(thead.getRule())){
			int level = CommonUtils.str2Int(thead.getLevel());
			CustomFile  cfile = getCustomFileByLevel(getParentsFile(invatation), level);
			String fileName =  cfile.getName();
			templist.add(fileName);
		}else if(Constants.RuleType.content.getName().equals(thead.getRule())){
			//TODO  内容解析
			String maches =  "";
			String keyword = thead.getKey();
			if(keyword == null|| "".equals(keyword)){
				keyword = thead.getTitle();
				if(keyword.equals("招标人")){
					System.out.println(key);
				}
			}
			if("table".equals(thead.getContentType())){
				maches = POIUtils.analysisTableString(ready2AnalyTables, keyword,thead);
			}else{
				
				if(thead.getDataType().equals(Constants.DataType.enumeration.getName())){
					List<Enumeration> es = thead.getEnumeration();
					String[] s =  new String[es.size()]; 
					int i =0 ;
					for(Enumeration e :es){
						s[i] = e.getValue();
						++i;
					}
					maches = POIUtils.analysisString(ready2AnalyParagraphs , s);
				}else if(CommonUtils.indexOf(keyword, new String[]{"是","以","为","应为"})!=-1){
					keyword = keyword.substring(0, keyword.length()-1);
					maches = POIUtils.analysisString(ready2AnalyParagraphs, keyword ,thead);
				}else{					
					maches = POIUtils.analysisString(ready2AnalyParagraphs, keyword, Constants.PATTERN1,thead);
				}
			}	
			templist.add(maches);
		}
	}
	private List<CustomFile> getDocsByName(CustomFile parnet){
		List <CustomFile> files = new ArrayList<CustomFile>();
		List <CustomFile> children = new ArrayList<CustomFile>();
		getChildrenFile(children, parnet);
		for (CustomFile customFile : children) {
			if(customFile.getName().endsWith(Constants.FileType.doc.getName())||
					customFile.getName().endsWith(Constants.FileType.docx.getName())){
				files.add(customFile);
				
			}
		}
		
		return  files;
	}
	private CustomFile getCustomFileByLevel(List<CustomFile> list , int level){
		if(!CommonUtils.isNull(list)){
			for (CustomFile customFile : list) {
				if(customFile.getLevel()==level){
					return customFile;
				}
			}
		}
		return null;
	}
	/**
	 * 递归 遍历 所有文件
	 * @param allFile
	 * @param file
	 * @param level
	 * @param parent
	 */
	private void listAllFile(List <CustomFile> allFile, File file,int level,CustomFile parent ){
		if(file.isDirectory()){			
			File [] files =  file.listFiles();
			if(files!=null&&files.length>0){
				List<CustomFile> children =  new ArrayList<CustomFile>(files.length);
				for (File child : files) {
					if(child.isHidden()){
						continue;
					}
					CustomFile customFile = new CustomFile();
					customFile.setLevel(level);
					customFile.setId(0);
					customFile.setName(child.getName());
					customFile.setAbsolutePath(child.getAbsolutePath());
					customFile.setParent(parent);
					allFile.add(customFile);
					children.add(customFile);
					if(child.isFile()){					
						customFile.setFolder(false);
						try {
							POIUtils.setCustomFileContent(customFile);
						} catch (NoSuchMethodException e) {
							System.out.println("path---->" + customFile.getAbsolutePath());
							e.printStackTrace();
						}
					}else{
						customFile.setFolder(true);
						listAllFile(allFile, child, level+1,customFile);
					}
				}
				parent.setChildren(children);
			}
		}	
		
	}
	private Integer getMapKeyByValue(Map<Integer,String> map,String v){
		Iterator<Entry<Integer, String>> it = map.entrySet().iterator();
		int key = -1;
		while (it.hasNext()) {
			Entry<Integer, String> item = it.next();
			if(item.getValue().equals(v)){
				key = (Integer) item.getKey();
				break;
			}
			
		}
		return key;
	}
	private List<CustomFile> getMainFolder(String key) {
		List<CustomFile> files = listAllCustomFile(root);
		List<CustomFile> ps = new ArrayList<CustomFile>();
		for (CustomFile customFile : files) {
			if(customFile.getName().equals(key)){
				ps.add(customFile.getParent());
			}
		}
		return ps;
	}
	private List<CustomFile> getParentsFile(CustomFile child){
		List<CustomFile> ps = new ArrayList<CustomFile>();
		CustomFile p =  child.getParent();
		if(p!=null){
			ps.add(p);
			getParent(ps, p);
		}
		return ps;
	}
	private void getParent(List<CustomFile>list ,CustomFile file){
		CustomFile p =  file.getParent();
		if(p!=null){
			list.add(p);
			getParent(list, p);
		}
	}
	@SuppressWarnings("unused")
	private List<CustomFile> getCustomFileByName(String name) {
		List<CustomFile> files = listAllCustomFile(root);
		List<CustomFile> ps = new ArrayList<CustomFile>();
		for (CustomFile customFile : files) {
			if(customFile.getName().equals(name)){
				ps.add(customFile);
			}
		}
		return ps;
	}
	private CustomFile getCustomFileByName(List<CustomFile> list, String name) {
		List<CustomFile> files = list;
		CustomFile ps = new CustomFile();
		for (CustomFile customFile : files) {
			if(name.equals(customFile.getName())){
				ps = customFile;
				break;
			}
		}
		return ps;
	}
	/**
	 * 
	 * @param file
	 * @param offset
	 * @return
	 */
	private List <CustomFile> getCustomFileByLevelOffset(CustomFile file,int offset){
		List <CustomFile> files = new ArrayList<CustomFile>();
		int level = file.getLevel()+ offset;
		if(offset==0){
			files.addAll(getBrotherFile(file));
		}else if(offset < 0){
			List<CustomFile> parents =  getParentsFile(file);
			for (CustomFile p : parents) {
				if(p.getLevel()==level){
					files.add(p);
				}
					
			}
		} else {
			List <CustomFile> children = new ArrayList<CustomFile>();
			getChildrenFile(children,file);
			if(!CommonUtils.isNull(children)){
				for (CustomFile customFile : children) {
					if(customFile.getLevel()==level){
						files.add(customFile);
					}
				}
			}
		}
		return  files;
	}
	
	private void getChildrenFile(List <CustomFile> files ,CustomFile p){
		List <CustomFile> children = p.getChildren();
		if(!CommonUtils.isNull(children)){
			//files.addAll(children);
			for (CustomFile c : children) {
				if(c != null){
					files.add(c);
					getChildrenFile(files,c);
				}
			}
		}
	}
	private List <CustomFile> getBrotherFile(CustomFile file){
		List <CustomFile> files = new ArrayList<CustomFile>();
		List<CustomFile> children = file.getParent().getChildren();
		for (CustomFile c : children) {
			if(!c.getName().equals(file.getName())){
				files.add(c);
			}
		}
		return  files;
	}
	public Map<Integer ,String > assembleData(String key,File excel){
		
		return null;
	}
	
	
}
