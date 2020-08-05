(Java poi Word模板数据替换)
## 需求说明
>在项目中遇到一个需要将报表数据导出到word的需求，报表数据包含了文字，图片和表格数据。

## 设计思路
 实际上在日常项目中，这种需求很常见，首先想到的思路就是可以事先定义好一个word模板,在模板中定义好变量，然后可以将报表数据对象的属性名称与变量名称对应，只要匹配上模板中的变量就将报表数据对象的值替换进去即可，思路有了，接下来开始细节的考虑。
 1. 首先是定义模板，可以在模板中用一对花括号将变量进行包裹，例如“{name}”, 其中的“{” 和“}”就是标记符号，当然至于使用花括号标记还是中括号来标记，最好是可以自定义，其中的“name” 就是报表数据对象的一个字段名称,例如：`private String name`。
 2. 需求中提到过要**支持图片**，通过研究poi 的api 文档发现它支持向word中插入图片，插入图片仅仅只需要一个图片的尺寸和一个图片地址，问题不大。
 3. 报表数据中还包含**表格**，报表中通常可能还不止一张表格（头大），也就是说在模板的不同位置需要插入不同的表格，其实可以在对象中设置不同的表格变量来解决，先分析表格数据通常会是一个List，List里面包含了一个记录每列数据的对象，可以考虑通过设置不同的标记符号来定义表格数据里面的变量，目的是为了与外层的数据对象的变量区分开。以下图为例说明
![在这里插入图片描述](https://img-blog.csdnimg.cn/20200804185516130.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L2JxcmZ2YXRq,size_16,color_FFFFFF,t_70#pic_center)
图中“{name}” 被定义为普通的文本类型变量，“{image}”被定义为一个图片类型变量，“{tableList1}”被定义为一个表格List变量，表格中用尖括号包裹的变量`<name>`被定义为表格数据对象的变量，那么问题来了，**在java 实体对象中如何区分是普通文本变量，还是图片类型变量，或者是表格类型变量了？？** 丝毫不慌，其实可以编写一个注解来标识字段的类型，下图代码示例说明：

```java
    @WordParams(type=WordParamsType.TEXT) 
	private String name;  //这个属性被标识为普通文本变量 对应模板中的{name}
	
	@WordParams(type=WordParamsType.FILE)
	private ImageInf image; //这个属性被标识为图片文件变量 对应模板中的{image}
	
	@WordParams(type=WordParamsType.LIST)
	private List<TableInfo> tableList1;//这个属性被标识为表格List变量  对应模板中的{tableList1}
```
可以看到，变量都可以用注解@WordParams 标识，通过对其设置不同的type就可以区分开来，如果是**图片变量**，可以看到这个字段image是一个对象属性**ImageInfo**,这个对象中维护了图片的尺寸和路径。如果是表格变量，那么这个字段就必须是一个集合属性List<?>上面代码中的“TableInfo” 指的是表格实体对象，也就是尖括号中的内容，如图所示：
![在这里插入图片描述](https://img-blog.csdnimg.cn/20200804192832378.png)
图中`<name> <age><address><telNo>` 对应的实体对象，如以下示例代码：

```java
    //注意 ： 这里的注解跟之前那个注解不同，这个注解是用来标识表格对象的字段的
    @WordTableParams //注意 ： 这里的注解跟之前那个注解不同
	private String name; //对应模板中的<name>
	
	@WordTableParams
	private String age;  //对应模板中的 <age>
	
	@WordTableParams
	private String address;  //对应模板中的 <address>
	
	@WordTableParams
	private String telNo;    //对应模板中的 <telNo>
```

> 注意 ： 这里的注解跟之前那个注解不同，这个注解是用来标识表格对象的字段的

## 扩展性
作为一个工具类，考虑到其可扩展性，程序在设计时预留的几个钩子方法，如下：



>  1.设置变量的前缀标记符号的，允许子类重写自定义，默认是 “{”
```java
/**
	 * 设置左边模板字符串
	 * 
	 * @param left
	 * @return
	 */
	public String setPreFix(String left) {
		// TODO Auto-generated method stub
		return left;
	}
```
> 2.设置变量的后缀标记符号,允许子类重写自定义，默认是 “}”

```java
/**
	 * 设置右边模板字符串
	 * 
	 * @param right
	 * @return
	 */
	public String setSuffix(String right) {
		// TODO Auto-generated method stub
		return right;
	}
```

> 3.设置表格变量的前缀标记符号，允许子类重写自定义，默认是“<”

```java
   /**
	 * 设置左标记
	 * @param left
	 * @return
	 */
	public String setTableSuffix(String left) {
		// TODO Auto-generated method stub
		return left;
	}
```

> 4.设置表格变量的后缀标记符号，允许子类重写自定义，默认是“>”

```java
  /**
     * 设置右标记
     * @param right
     * @return
     */
	public String setTablePreFix(String right) {
		// TODO Auto-generated method stub
		return right;
	}
```

> 5.从文件中提取段落集，poi 操作Word文件读取内容，是从文件的指定部分比如：从文本段落、表格、页眉、页脚中提取段落进行读取，目前只实现了从文本段落和表格中提取段落，如果需要从其他部分提取，可以允许子类从`document`中获取，然后添加到`findList` 段落集合中。

```java
/**
	 * 获取段落从整个文件中
	 * 
	 * @param document
	 * @param arrayList
	 * @return
	 */
	public List<DocInfo> findDocInfByDocument(XWPFDocument document, List<DocInfo> findList) {
	}
```

> 6.程序中对不同类型的变量设计了不同的处理流程，目前只支持文本、图片、表格类型的处理，如果在满足不了需求的情况下，可以通过重写`otherReplaceHander`方法扩展其他类型的替换处理流程

```java
/**
      * 其他变量类型替换方案
	 * @param document 
    * @param bean
    * @param concatText
	 * @param doc 
	 * @param xWPFRunList 
    */
	public void otherReplaceHander(XWPFDocument document, ParamsBean bean, String concatText, DocInfo doc, List<XWPFRun> xWPFRunList) {
	}
```
> 7.导出前的钩子，在程序执行到的导出前，如果还需要一些额外的功能要实现，可以通过重写`beforWriterHandle`这个方法进行扩展

```java
/**
	 * 导出前处理方法
	 * 
	 * @param pojoParamList
	 * @param document
	 */
	public void beforWriterHandle(XWPFDocument document, List<ParamsBean> pojoParamList) {
		// TODO Auto-generated method stub
	}
```

> 8.文件导后的钩子，在程序执行导出后，如果还需要一些额外的功能要实现，可以通过重写`afterWriterHandle`这个方法进行扩展

```java
/**
	 * 导出之后处理方法钩子
	 * 
	 * @param document
	 * @param t
	 */
	public void afterWriterHandle(XWPFDocument document, BaseWordTemp t) {
		// TODO Auto-generated method stub
	}
```

## 使用
前面啰嗦的了半天，终于要说重点了，使用这个工具可分为以下四个步骤：
 1. 制作模板
 相信前面的内容看完，大概已经知道模板如何制作了，这里就不再赘述了，直接丢过来一个示例模板：[word模板示例](https://github.com/bqrfvatj/project/blob/office-temp-tool/word%E7%A4%BA%E4%BE%8B%E6%A8%A1%E6%9D%BF.docx)
 2. 配置开发环境
  把[Word模板工具类文档](https://github.com/bqrfvatj/project/tree/office-temp-tool/Word%E6%A8%A1%E6%9D%BF%E5%B7%A5%E5%85%B7%E7%B1%BB%E6%96%87%E6%A1%A3)中的jar包和**mavn_install_cfg.bat** （安装脚本）以及脚本配置文件（**config.bcfg**），放到你本地任意一个目录下，前提是你的电脑已经配置好了maven 的环境变量。
 ![在这里插入图片描述](https://img-blog.csdnimg.cn/20200805113615693.png)
修改config.bcfg，填写jar包的maven配置信息；然后双击安装脚本，不出意外的话，jar包已经安装到你maven的本地仓库了，复制提示框中的maven配置信息，添加到你的项目pom文件中去，就算配置好开发环境了。
 
 4. 构建测试数据
 

> 3.1 构建报表数据实体类，该类必须继承`BaseWordTemp`

```java
@Setter
@Getter
public class ReportDataInfo extends BaseWordTemp{
	
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
	private List<UserInfo> tableList2;
	
}
```

> 3.2 创建第二张表结构，tableList1字段的类型TableInfo实体类，字段的注解与报表数据实体类字段上的注解不同

```java
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
```
> 3.3 创建第二张表结构，tableList2字段的类型UserInfo实体类，字段的注解与报表数据实体类字段上的注解不同


```java
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
```

> 3.4 构造表格测试数据

```java
private void buildListData(ReportDataInfo bodyInfo) {
		// TODO Auto-generated method stub
		//第一张表的模拟数据
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
		
		//第二张表的模拟数据
		List<UserInfo> tableList2 = new ArrayList<UserInfo>();
		for(int i=0;i<10;i++) {
			UserInfo tableInfo = new UserInfo();
			tableInfo.setWorker("CEO");
			tableInfo.setLike("女");
			tableInfo.setWorkYear(5+i);
			tableInfo.setSex("男");
			tableList2.add(tableInfo);
		}
		bodyInfo.setTableList2(tableList2);
	}
```


 4. 执行数据替换并导出到新的目录
  

```java
public static void main(String[] args) {
		WordTester wordTester = new WordTester();
		ReportDataInfo reportDataInfo = new ReportDataInfo();
		reportDataInfo.setYear("2019");//文本
		reportDataInfo.setMonth("10");//文本
		reportDataInfo.setImage1(new ImageInf(200,200,"D:\\Test\\file\\timg.jpg")); //图片
		reportDataInfo.setImage2(new ImageInf(300,200,"D:\\Test\\file\\ss.jpg"));//图片
		reportDataInfo.setTempPath("D:\\Test\\file\\word示例模板.docx");    //word 模板路径
		reportDataInfo.setOutPath("D:\\Test\\file\\testFolder\\testResult.docx"); //替换后word的导出路径
		//构造表格测试数据
		wordTester.buildListData(reportDataInfo);
		//word 替换工具的实例，该类继承自AbstractWordTemple，具体细节见《Word模板工具类API.docx》
		TableLoopReplaceHandle wordUtil = new TableLoopReplaceHandle();
		//调用主方法执行报表数据导出到word
		wordUtil.findLabelAndReplace(reportDataInfo);
	}
```

## 结尾
个人建议啊，在技术选型的时候要有自己的主见，不能人云亦云，就好比我在写这个工具的时候，大家都说poi不好用，看到@象话 这位女博主就提到过，可是我就坚持用poi,研究了2小时，确实不好用，哈哈开个玩笑。

> 如果有小伙伴对源码感兴趣的话，在这里分享[源码地址](https://github.com/bqrfvatj/project/tree/office-temp-tool)给大家，欢迎来找茬！！

