package com.word2Excel.bean;
/**
 * 
 * Í¶±ê
 * @author Administrator
 *
 */
public class Bid {
	private String bidName = "";
	private String prices = "";
	
	public Bid(){
		
	}
	public Bid(String bidName, String prices) {
		super();
		this.bidName = bidName;
		this.prices = prices;
	}
	public String getBidName() {
		return bidName;
	}
	public void setBidName(String bidName) {
		this.bidName = bidName;
	}
	public String getPrices() {
		return prices;
	}
	public void setPrices(String prices) {
		this.prices = prices;
	} 
	
}
