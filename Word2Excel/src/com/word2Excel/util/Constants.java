package com.word2Excel.util;

public class Constants {	
	/**
	 *   filter 
	 */
	public  final static String TYPE_INVITATION_FOR_BIDS = "01-招标文件";
	public  final static String TYPE_TENDER = "02-投标文件";   
	
	
	public  final static String TYPE_BUSINESS = "商务部分";
	
	/**
	 * 
	 * 分割符
	 *
	 */
	public  static enum Splitor{
		colon(":"),
		colon_zh("："),
		comma(","),
		comma_zh("，"),
		full_stop("。");   //句号
		
		private String name;
		Splitor(String name){
			this.name = name;
		}
		public String getName() {
			return name;
		}
		public void setName(String name) {
			this.name = name;
		}
		
	}
	public static enum FileType{
		doc("doc"),
		docx("docx"),
		txt("txt"),
		excel("xlsx");
		
		private String name;
		FileType( String name) {
			this.name = name;
		}
		public String getName() {
			return name;
		}
		public void setName(String name) {
			this.name = name;
		}
	}
	
	public static enum RuleType{
		folder("folder"),
		content("content");
		
		private String name;
		RuleType( String name) {
			this.name = name;
		}
		public String getName() {
			return name;
		}
		public void setName(String name) {
			this.name = name;
		}
	}
	
	public final static String[] RULES_TYPE ={
			"folder","content"
			}; 
	
	public static enum Regex{
		number("^[0-9]+");
		private String name;
		Regex( String name) {
			this.name = name;
		}
		public String getName() {
			return name;
		}
		public void setName(String name) {
			this.name = name;
		}
		
	}
	public  final static String  PATTERN  = "^[\\(（][^\\(（]+[\\)）]$"; //匹配 " () "
	public  final static String  PATTERN1  = "\\S+(:){1,}\\S+|\\S+(：){1,}\\S+"; //匹配 " **:** "
	
	
	
	
}
