package org.ubfs.word.temp.junit;

import org.ubfs.word.temp.annoation.WordTableParams;

import lombok.Data;

@Data
public class UserInfo {
	@WordTableParams
	private String worker;
	@WordTableParams
	private String like;
	@WordTableParams
	private Integer workYear;
	@WordTableParams
	private String sex;

}
