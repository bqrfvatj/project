package org.ubfs.word.temp.junit;

import org.ubfs.word.temp.annoation.WordTableParams;

import lombok.Data;

@Data
public class TableInfo {
	
	@WordTableParams //注意 ： 这里的注解跟之前那个注解不同
	private String name; 
	
	@WordTableParams
	private String age;
	
	@WordTableParams
	private String address;
	
	@WordTableParams
	private String telNo;
	
}
