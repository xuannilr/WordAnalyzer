package com.word2Excel.util;

public class Constants {	
	/**
	 *   filter 
	 */
	public  final static String TYPE_INVITATION_FOR_BIDS = "01-�б��ļ�";
	public  final static String TYPE_TENDER = "02-Ͷ���ļ�";   
	
	
	public  final static String TYPE_BUSINESS = "���񲿷�";
	
	/**
	 * 
	 * �ָ��
	 *
	 */
	public  static enum Splitor{
		colon(":"),
		comma(","),
		colon_zh("��");
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
	
	public  final static String  PATTERN  = "^[\\(��][^\\(��]+[\\)��]$"; //ƥ�� " () "
	public  final static String  PATTERN1  = "\\S+(:){1,}\\S+|\\S+(��){1,}\\S+"; //ƥ�� " **:** "
	
	
	
	
}