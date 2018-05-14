package com.word2Excel.bean.vo;

import java.util.ArrayList;
import java.util.List;

public class ProjectItem {
	private List<Group> groups = new ArrayList<Group>();

	public List<Group> getGroups() {
		return groups;
	}

	public void setGroups(List<Group> groups) {
		this.groups = groups;
	}
	
}
