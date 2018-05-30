package com.word2Excel.bean.vo;


import java.util.ArrayList;
import java.util.List;

import org.dom4j.Element;

import com.word2Excel.util.CommonUtils;

public class Thead {
	private String rule ;
	private String dataType;
	private String key;
	private String title;
	private String contentType;
	private String level;
	private String sukey ;
	private String direction;
	private String unit;
	List<Enumeration>  enumeration;
	
	public List<Enumeration> getEnumeration() {
		return enumeration;
	}
	public void setEnumeration(List<Enumeration> enumeration) {
		this.enumeration = enumeration;
	}
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
	
	public String getUnit() {
		return unit;
	}
	public void setUnit(String unit) {
		this.unit = unit;
	}
	public Thead(Element ele){
		this();
		if(ele!=null){
			this.rule = ele.attributeValue("rule");
			this.contentType = ele.attributeValue("content-type");
			this.title = ele.attributeValue("title");
			this.key = ele.attributeValue("key");
			this.dataType = ele.attributeValue("type");
			this.level = ele.attributeValue("level");
			this.sukey= ele.attributeValue("sukey");
			this.direction = ele.attributeValue("direction");
			this.unit = ele.attributeValue("unit");
			List <Element> eumertations =  ele.elements("enumeration");
			if(!CommonUtils.isNull(eumertations)){
				this.enumeration = new ArrayList<Enumeration>();
				for(Element eumer :eumertations){
					this.enumeration.add(new Enumeration(eumer));
				}
			}
			
		}
	}
	public Thead(){
		
	}
}
