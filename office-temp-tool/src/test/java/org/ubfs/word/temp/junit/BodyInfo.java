package org.ubfs.word.temp.junit;

import java.util.List;

import org.ubfs.word.temp.annoation.WordParams;
import org.ubfs.word.temp.beans.BaseWordTemp;
import org.ubfs.word.temp.beans.ImageInf;
import org.ubfs.word.temp.constant.WordParamsType;

import lombok.Getter;
import lombok.Setter;

@Setter
@Getter
public class BodyInfo extends BaseWordTemp{
	
	@WordParams
	private String year;
	@WordParams
	private String month;
	@WordParams
	private String title;
	
	@WordParams(type=WordParamsType.FILE)
	private ImageInf image1;
	
	@WordParams(type=WordParamsType.FILE)
	private ImageInf image2;
	
	
	@WordParams(type=WordParamsType.LIST)
	private List<TableInfo> tableList1;
	
	@WordParams(type=WordParamsType.LIST)
	private List<UserInfo> tableList;
	

}
