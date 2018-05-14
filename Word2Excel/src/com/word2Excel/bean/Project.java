package com.word2Excel.bean;

import java.util.ArrayList;
import java.util.List;

import com.word2Excel.bean.vo.Group;

/**
 * ����
 * @author Administrator
 *
 */
public class Project {
	private String pSequence ;  	//�б���   
	private String pDate;		    //��������		
	private String pCorporation;	//��Ŀ��˾
	private String pName;  			//������+�����+����
	
	private List<Bid> bids = new ArrayList<Bid>();
	
	
	public Project(){
		
	}
	
	public Project(String pSequence, String pDate, String pCorporation, String pName) {
		super();
		this.pSequence = pSequence;
		this.pDate = pDate;
		this.pCorporation = pCorporation;
		this.pName = pName;
	}
	
	public String getpSequence() {
		return pSequence;
	}
	
	public void setpSequence(String pSequence) {
		this.pSequence = pSequence;
	}
	public String getpDate() {
		return pDate;
	}
	public void setpDate(String pDate) {
		this.pDate = pDate;
	}
	public String getpCorporation() {
		return pCorporation;
	}
	public void setpCorporation(String pCorporation) {
		this.pCorporation = pCorporation;
	}
	public String getpName() {
		return pName;
	}
	public void setpName(String pName) {
		this.pName = pName;
	}

	public List<Bid> getBids() {
		return bids;
	}

	public void setBids(List<Bid> bids) {
		this.bids = bids;
	}
}
