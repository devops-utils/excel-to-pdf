# excel-to-pdf
An easy way to convert Excel 2003 and Excel 2007 to PDF by Java code based on Apache POI and itext7

### 使用SDK
```java
IExcel2PDF excel2PdfTool = EPFactory.getEP("case6.xlsx", "output1.pdf", System.getProperty("user.dir") + "/doc/font/SimHei.TTF");
if(excel2PdfTool != null) {
    excel2PdfTool.convert();
}
```

### maven
### Maven 引用方式
```xml
<dependency>
    <groupId>com.github.zhangchunsheng</groupId>
    <artifactId>excel-to-pdf</artifactId>
    <version>1.0.2</version>
</dependency>
```

```
https://github.com/itext/itext7
https://www.tutorialspoint.com/itext/itext_text_annotation.htm
```

<table border="0">
	<tbody>
		<tr>
			<td align="center" valign="middle">
				<a href="https://url.cn/5jVTRwI" target="_blank">
					<!--<img height="120" src="https://wx4.sinaimg.cn/mw690/46b94231ly1ge0pvo2necj209l05kq3c.jpg">-->
					<img height="120" src="https://ride-group.gitee.io/amapjava/images/tencent.jpeg">
				</a>
			</td>
			<td align="right" valign="middle">
				<!--<img height="120" src="https://wx2.sinaimg.cn/mw690/46b94231ly1ge0po9ko70j20fk0fkjsc.jpg">-->
				<img height="120" src="https://ride-group.gitee.io/amapjava/images/fenxiang.jpeg">
			</td>
			<td align="center" valign="middle">
				<a href="https://www.vultr.com/?ref=8546025-6G" target="_blank">
					<!--<img height="120" src="https://wx3.sinaimg.cn/mw1024/46b94231ly1ge0p76k64bj206o06owev.jpg">-->
					<img height="120" src="https://ride-group.gitee.io/amapjava/images/vultr.jpeg">
				</a>
			</td>
			<td align="center" valign="middle">
				<a href="https://www.aliyun.com/minisite/goods?userCode=tewwu0c8" target="_blank">
					<!--<img height="120" src="https://img.alicdn.com/tfs/TB1Gc3zmAL0gK0jSZFxXXXWHVXa-259-194.jpg">-->
					<img height="120" src="https://ride-group.gitee.io/amapjava/images/aliyun.jpeg">
				</a>
			</td>
		</tr>
	</tbody>
</table>

## 捐助 donate

<table border="0">
	<tbody>
	    <tr>
	        <td>支付宝</td>
	        <td>微信</td>
	    </tr>
		<tr>
			<td align="left" valign="middle">
                <!--<img height="120" src="https://wx4.sinaimg.cn/mw690/46b94231ly1ge0okee0fej20ec0e6gp3.jpg">-->
                <img height="120" src="https://ride-group.gitee.io/amapjava/images/alipay.jpeg">
			</td>
			<td align="center" valign="middle">
				<!--<img height="120" src="https://wx4.sinaimg.cn/mw690/46b94231ly1ge0okecldyj20e80e8n0c.jpg">-->
				<img height="120" src="https://ride-group.gitee.io/amapjava/images/wechat.jpeg">
			</td>
		</tr>
	</tbody>
</table>