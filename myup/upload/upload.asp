<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
OPTION EXPLICIT
Server.ScriptTimeOut=5000
%>
<!--#include FILE="../UpLoadClass.asp"-->
<!--#include file="../help/Code2Info.en.asp"-->
<%
dim request2,formPath,formName,intCount,intTemp
'建立上传对象
set request2=new UpLoadClass
	'设置文件允许的附件类型为gif/jpg/rar/zip/默认值："gif/jpg" 特征值：""（空） 表示文件类型不受限制
	request2.FileType=""

	'设置服务器文件保存路径
	request2.SavePath="UpLoadFile/"

	'设置字符集
	request2.Charset="UTF-8"

	request2.AutoSave=1

	'打开对象
	request2.Open()
%>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>上传成功</title>
</head>
<body>
<%
	'显示类版本
	Response.Write("<p>程序版本："&request2.Version&"</p>")

	'显示字符集
	Response.Write("<p id=""Title"">字 符 集："&request2.Charset&"</p>")

	'显示邮件标题
	Response.Write("<p>邮件标题："&request2.Form("strTitle")&"</p>")

	'显示邮件内容
	Response.Write("<p>邮件内容："&request2.Form("strContent")&"</p>")

	'----列出所有上传了的文件开始----
	intCount=0
	for intTemp=1 to Ubound(request2.FileItem)

		'获取表单文件控件名称，注意FileItem下标从1开始
		formName=request2.FileItem(intTemp)

		'显示文件域
		Response.Write( "<p>"&intTemp&"、文件域["&formName&"]：")

		'显示源文件路径与文件名
		Response.Write( "<br />"&request2.form(formName&"_Path")&request2.form(formName&"_Name"))

		'显示文件大小（字节数）
		Response.Write( "("&request2.form(formName&"_Size")&") => ")

		'显示目标文件路径与文件名
		Response.Write(request2.SavePath&request2.Form(formName)&"</p>")

		if request2.form(formName&"_Err")=0 then intCount=intCount+1

		'显示文件保存状态
		'Err2Info代码转信息，见../Code2Info.en.asp
		Response.Write("<br />"&Err2Info(request2.form(formName&"_Err"))&"</p>")

	next
	'----列出所有上传了的文件结束----

	Response.Write( "<p>共 "&intCount&" 个文件上传成功! </p>")

	Response.Write( "<p>[<a href=""javascript:history.back();"">返回</a>]</p>")
	%>
</body>
</html>
<%
'释放上传对象
set request2=nothing
%>