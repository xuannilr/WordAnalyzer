package com.word2Excel.bean.vo;

import org.dom4j.Element;

public class Enumeration {
	private String value;
	
	Enumeration(){}
	Enumeration(Element e){
		if(e!=null){
			this.setValue(e.getText());
		}
	}
	public String getValue() {
		return value;
	}
	public void setValue(String value) {
		this.value = value;
	}
}
