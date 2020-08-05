package org.ubfs.word.temp.beans;

import lombok.Data;

@Data
public class ParamsBean {
	
	private int type;//类型
	private String field;//预备替换的字段
	private Object value;//预备替换的内容
	
	private BaseWordTemp baseWordTemp;
	
	public BaseWordTemp getBaseWordTemp() {
		return baseWordTemp;
	}
	public void setBaseWordTemp(BaseWordTemp baseWordTemp) {
		this.baseWordTemp = baseWordTemp;
	}
	@Override
	public String toString() {
		return "ParamsBean [type=" + type + ", field=" + field + ", value=" + value + "]";
	}
	
	
	
	

}
