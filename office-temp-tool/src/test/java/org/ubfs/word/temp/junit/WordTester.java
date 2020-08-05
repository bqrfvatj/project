package org.ubfs.word.temp.junit;

import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.junit.Test;
import org.ubfs.word.temp.beans.ImageInf;
import org.ubfs.word.temp.service.imple.TableLoopReplaceHandle;

public class WordTester {
	
	@Test
	public void testMethod() {
		BodyInfo bodyInfo =  new BodyInfo();
		bodyInfo.setYear("2019");
		bodyInfo.setMonth("10");
		bodyInfo.setImage1(new ImageInf(200,200,"D:\\Test\\file\\timg.jpg"));
		bodyInfo.setImage2(new ImageInf(300,200,"D:\\Test\\file\\ss.jpg"));
		bodyInfo.setTempPath("D:\\Test\\file\\test3.docx");
		bodyInfo.setOutPath("D:\\Test\\file\\testFolder\\testResult.docx");
		buildListData(bodyInfo);
		TableLoopReplaceHandle wordUtil = new TableLoopReplaceHandle();
		wordUtil.findLabelAndReplace(bodyInfo);
	}
	
	private void buildListData(BodyInfo bodyInfo) {
		// TODO Auto-generated method stub
		List<TableInfo> tableList1 = new ArrayList<TableInfo>();
		for(int i=0;i<10;i++) {
			TableInfo tableInfo = new TableInfo();
			tableInfo.setName("张" + i);
			tableInfo.setAddress("福田保税区");
			tableInfo.setAge("18");
			tableInfo.setTelNo("1888885"+i+"278");
			tableList1.add(tableInfo);
		}
		bodyInfo.setTableList1(tableList1);
		
		
		List<UserInfo> tableList2 = new ArrayList<UserInfo>();
		for(int i=0;i<10;i++) {
			UserInfo tableInfo = new UserInfo();
			tableInfo.setWorker("CEO");
			tableInfo.setLike("女");
			tableInfo.setWorkYear(5+i);
			tableInfo.setSex("男");
			tableList2.add(tableInfo);
		}
		bodyInfo.setTableList(tableList2);
	}

	@Test
	public void checkTextFormat() {
		String concatText = "}{image3}";
		Pattern r1 = Pattern.compile("(\\{[^\\}]*\\})");
		Pattern r2 = Pattern.compile("\\{");
		Matcher m1 = r1.matcher(concatText);
		Matcher m2 = r2.matcher(concatText);
		int successCount = 0;
		int count = 0;
		while(m1.find()) {
			successCount ++ ;
		}
		while(m2.find()) {
			count ++ ;
		}
		boolean is = count > successCount ? false : true;
		
		System.out.println(is);
	}
	@Test
	public void testMethod2() {
	  String ss = "}{112} {222";
      Pattern r = Pattern.compile("(\\{[^\\}]*\\})");
      
      Pattern r2 = Pattern.compile("\\{");
    
      Matcher m = r2.matcher(ss);
      Matcher m2 = r.matcher(ss);
      int length = 0;
      while(m.find()) {
    	  length ++ ;
      }
      while(m2.find()) {
    	  System.out.println(m2.group());
      }
      
      System.out.println(length);
	}

}
