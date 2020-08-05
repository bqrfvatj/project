package org.ubfs.word.temp.junit;

import java.util.List;

import org.ubfs.word.temp.annoation.WordParams;
import org.ubfs.word.temp.beans.BaseWordTemp;
import org.ubfs.word.temp.beans.ImageInf;
import org.ubfs.word.temp.constant.WordParamsType;

import lombok.Data;

@Data
public class ApiInfo extends BaseWordTemp{
	
	@WordParams(type=WordParamsType.TEXT) 
	private String name;  //这个属性被标识为普通文本变量
	
	@WordParams(type=WordParamsType.FILE)
	private ImageInf image; //这个属性被标识为图片文件变量
	
	@WordParams(type=WordParamsType.LIST)
	private List<TableInfo> tableList1;//这个属性被标识为表格List变量

}
