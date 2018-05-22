package com.word2Excel.bean.vo;

import org.dom4j.Element;

public class Thead {
	String rule ;
	String dataType;
	String key;
	String title;
	String property;
	String contentType;
	String level;
	String sukey ;
	String direction;
	
	
	public String getLevel() {
		return level;
	}
	public void setLevel(String level) {
		this.level = level;
	}
	public String getRule() {
		return rule;
	}
	public void setRule(String rule) {
		this.rule = rule;
	}
	public String getDataType() {
		return dataType;
	}
	public void setDataType(String dataType) {
		this.dataType = dataType;
	}
	public String getKey() {
		return key;
	}
	public void setKey(String key) {
		this.key = key;
	}
	public String getTitle() {
		return title;
	}
	public void setTitle(String title) {
		this.title = title;
	}
	public String getProperty() {
		return property;
	}
	public void setProperty(String property) {
		this.property = property;
	}
	public String getContentType() {
		return contentType;
	}
	public void setContentType(String contentType) {
		this.contentType = contentType;
	}
	
	public String getSukey() {
		return sukey;
	}
	public void setSukey(String sukey) {
		this.sukey = sukey;
	}
	
	public String getDirection() {
		return direction;
	}
	public void setDirection(String direction) {
		this.direction = direction;
	}
	public Thead(Element ele){
		this();
		if(ele!=null){
			this.contentType = ele.attributeValue("content-type");
			this.key = ele.attributeValue("key");
			this.property = ele.attributeValue("property");
			this.rule = ele.attributeValue("rule");
			this.title = ele.attributeValue("title");
			this.dataType = ele.attributeValue("dataType");
			this.level = ele.attributeValue("level");
			this.sukey= ele.attributeValue("sukey");
			this.direction = ele.attributeValue("direction");
		}
	}
	public Thead(){
		
	}
}
