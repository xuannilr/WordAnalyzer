package com.word2Excel.bean.vo;

import java.util.ArrayList;
import java.util.List;

import org.dom4j.Element;

import com.word2Excel.util.CommonUtils;

public class Group {
	String key;
	String className;
	List <Thead> theads = new ArrayList<Thead>();
	
	public String getKey() {
		return key;
	}
	public void setKey(String key) {
		this.key = key;
	}
	public String getClassName() {
		return className;
	}
	public void setClassName(String className) {
		this.className = className;
	}
	public List<Thead> getTheads() {
		return theads;
	}
	public void setTheads(List<Thead> theads) {
		this.theads = theads;
	}
	
	public Group(Element ele){
		this();
		if(ele!=null){
			this.className = ele.attributeValue("class");
			this.key = ele.attributeValue("key");
			@SuppressWarnings("unchecked")
			List<Element> ths = ele.elements("th");
			if(!CommonUtils.isNull(ths)){
				for (Element element : ths) {
					this.theads.add(new Thead(element));
				}
			}
			

		}
	}
	public Group(){
		
	}
	
}
