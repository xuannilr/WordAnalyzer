package com.word2Excel.bean;

import java.util.List;
import java.util.Map;

public class CustomFile  {
	
	private int level;
	private int id;
	private String name;
	private String absolutePath;
	private CustomFile parent;
	private boolean isFolder;
	private List<CustomFile> children ;
	private List<String> paragrathsText ;
	private List <Map<String,String>> tablesParagraphsText;

	public CustomFile() {
		super();
	}
	
	public boolean isFolder() {
		return isFolder;
	}

	public void setFolder(boolean isFolder) {
		this.isFolder = isFolder;
	}

	public int getLevel() {
		return level;
	}
	public void setLevel(int level) {
		this.level = level;
	}
	public int getId() {
		return id;
	}
	public void setId(int id) {
		this.id = id;
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	
	public String getAbsolutePath() {
		return absolutePath;
	}
	public void setAbsolutePath(String absolutePath) {
		this.absolutePath = absolutePath;
	}
	public CustomFile getParent() {
		return parent;
	}
	public void setParent(CustomFile parent) {
		this.parent = parent;
	}
	public List<CustomFile> getChildren() {
		return children;
	}
	public void setChildren(List<CustomFile> children) {
		this.children = children;
	}

	public List<String> getParagrathsText() {
		return paragrathsText;
	}

	public void setParagrathsText(List<String> paragrathsText) {
		this.paragrathsText = paragrathsText;
	}

	public List<Map<String, String>> getTablesParagraphsText() {
		return tablesParagraphsText;
	}

	public void setTablesParagraphsText(List<Map<String,String>> tablesParagraphsText) {
		this.tablesParagraphsText = tablesParagraphsText;
	}
	
}
