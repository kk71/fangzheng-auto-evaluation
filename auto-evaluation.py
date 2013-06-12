#! /usr/bin/env python3
# -*- coding: utf-8 -*-

author_info='''
适用于“正方教学管理系统”的自动评教工具
通过测试学校：浙工大之江学院
作者：kk（kkblog.org，weibo.com/billfk）
'''


import sys,os
from datetime import datetime

from urllib.request import urlopen,install_opener,build_opener,Request,HTTPHandler,HTTPCookieProcessor
from http.cookiejar import CookieJar
from urllib.parse import urlencode
import re




#urls(仅适用于浙工大之江学院:P)
site_urls={
	"domain-url":"http://172.16.253.30/",#该url后将续接session代码，然后可以续接主页面或评教页面。
	"login":None,#登入点，等待获得session代码之后动态赋值
}

#========================================================
#登录成功的标志
login_sign="安全退出"


#========================================================
#找出主页中的评教课程区域
course_list_area_re=re.compile(r"教学质量评价.*信息维护")

#在区域中找出每个课程的anchor tag
course_link_re=re.compile(r'<a href=".*?".*?>.*?</a>')#return every tags of course

#从课程的tag中获得课程名和评教页面的相对url
link_re=re.compile(r'".*?"')#get the link to the course
course_name_re=re.compile(r">.*?</a>")#get current course name


#========================================================
#每个学科的网页中的评价form的tag
course_form_post_re=re.compile(r'<form name="Form1" method="post" action=".*?".*?>')

#找出那一长串的viewstate的值
viewstate_re=re.compile(r'name="__VIEWSTATE" value=".*?"')

#找出目前正在评的科目名（用于验证）和代号（用于返回post）的tag
#re.S让“.”能跨行匹配
current_course_area_re=re.compile(r'评价课程名称：<select.*?<option selected="selected" value=".*?">.*?</option>.*?</select>',re.S)
current_course_tag_re=re.compile(r'<option selected="selected" value=".*?">.*?</option>')

#获得考察指标
index_re=re.compile(r'<select name="DataGrid1:.*?" id="DataGrid1.*?">')

#获得课程种类
current_category_area_re=re.compile(r'<select name="DropDownList1".*?</select>',re.S)
current_category_tag_re=re.compile(r'<option selected="selected" value=".*?">.*?</option>')

#评价值
value_excellent="优秀".encode("gbk")
value_good="良好".encode("gbk")

#所有评价完成后，任意一个评价页面上必然出现以下标志
evaluation_finished_sign=" 提  交 "



#=========================================================
#prepare for cookie
cookie_support= HTTPCookieProcessor(CookieJar())
opener = build_opener(cookie_support, HTTPHandler)
install_opener(opener)


#获得服务器自动分配的session值
now_url=urlopen(site_urls["domain-url"]).url
site_urls["domain-url"]=re.compile(site_urls["domain-url"]+r".*?/").findall(now_url)[0]
site_urls["login"]=site_urls["domain-url"]+"default2.aspx"



#=========================================================
def login(login_info):
	'''
登录用户并获得登录成功的页面内容（用于取得课程列表）和登录后跳转的页面url（用于放入获取评教页面的headers里）
return:
	成功：tuple:第一个为url，第二个为页面内容（已gbk解码）
	失败：False
'''
	headers={
		"User-Agent":"Mozilla/5.0 (X11; Linux i686) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.97 Safari/537.11",
	}
	req=Request(url=site_urls["login"],data=urlencode(login_info).encode("utf-8"),headers=headers)
	try:
		
		urlobj=urlopen(req)
		content=urlobj.read().decode("gbk")#gbk decode
		url=urlobj.url
	except:
		return False

	if content.find(login_sign)!=-1:
		return (url,content)

	return False	



#=========================================================
def get_course_list(page):
	'''
'''
	course_area=course_list_area_re.findall(page)
	if course_area==[]:return False
	
	courses=dict()
	for tag in course_link_re.findall(course_area[0]):
		try:
			link=link_re.findall(tag)[0][1:-1]
			cname=course_name_re.findall(tag)[0][1:-4]
			courses.update({link:cname})
		except:
			return False
	return courses	



#=========================================================
def set_evaluation(course,url,main_page_url):
	'''
评教。
in:
	course(str):课程名称
	url(str):从主页读出的相对地址。
	main_page_url(str):主页的绝对url，用于放入headers

return:
	False：失败。
	True:全部评教完毕，且已提交。
'''

	#获得评教页面
	headers={
		"Referer":main_page_url,#
		"User-Agent":"Mozilla/5.0 (X11; Linux i686) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.97 Safari/537.11",
	}
	req=Request(url=site_urls["domain-url"]+url,headers=headers)
	try:
		content=urlopen(req).read().decode("gbk")#gbk解码
	except:
		return False
	

	#即将发送的表单
	#注意，该表单可能重复使用。当提交时用Button1即“保存当前课程的评教”时，
	#	提交用Button2即“当所有课程的评教都结束了提交所有保存过的信息”。后者为一次性操作。
	save_evaluation={
				"__EVENTTARGET":"",
				"__EVENTARGUMENT":"",
				"__VIEWSTATE":None,#viewstate,等待从页面读入
				"pjkc":None,#科目代码，等待从页面读入
				"DropDownList1":None,#科目种类，等待读入
				"txt1":"",
				"TextBox1":"40",
				"pjxx":"",#默认无留言
				"Button1":" 保  存 ".encode("gbk"),
	}


	#获得viewstate
	try:
		viewstate=viewstate_re.findall(content)[0][26:-1]#[26:-1]是从获得的内容里取出viewstate
	except:
		return False#can't find viewstate means got the wrong page
	save_evaluation.update({"__VIEWSTATE":viewstate})


	#获得种类代码
	category_area=current_category_area_re.findall(content)[0]
	category_tag=current_category_tag_re.findall(category_area)[0]
	category_val=re.compile(r'".*?"').findall(category_tag)[1][1:-1]
	save_evaluation.update({"DropDownList1":category_val})


	#获得当前评测的课程名和代码（course_name和course_val）
	try:
		course_area=current_course_area_re.findall(content)[0]
		course_tag=current_course_tag_re.findall(course_area)[0]

		course_val_re=re.compile(r'".*?"')
		course_val=course_val_re.findall(course_tag)[1][1:-1]

		course_name_re=re.compile(r'>.*?</option>')
		course_name=course_name_re.findall(course_tag)[0][1:-9]
	except:
		return False
	if course.strip().lower()!=course_name.strip().lower():return False
	save_evaluation.update({"pjkc":course_val})
	

	#获得评测指标。
	indices=index_re.findall(content)
	for n in indices:
		n=link_re.findall(n)[0][1:-1]
		save_evaluation.update({n:value_excellent})
	save_evaluation.update({n:value_good})
	

	#获得发送url
	try:
		post_to=course_form_post_re.findall(content)[0]
		tmp_re=re.compile(r'".*?"')
		post_to=tmp_re.findall(post_to)[2][1:-1]
	except:
		return False


	#发送
	headers.update({"Referer":site_urls["domain-url"]+url})
	try:
		req=Request(url=site_urls["domain-url"]+post_to,
				data=urlencode(save_evaluation).encode("gbk"),
				headers=headers)
		content=urlopen(req).read().decode("gbk")#gbk decode
	except:
		return False

	if content.find(evaluation_finished_sign)==-1:return

	#判别是否收到了“评测全部完成”标志
	#若收到了，则可以提交所有评测了。
	save_evaluation.pop("Button1")
	save_evaluation.update({"Button2":" 提  交 ".encode("gbk")})
	try:
		req=Request(url=site_urls["domain-url"]+post_to,
				data=urlencode(save_evaluation).encode("gbk"),
				headers=headers)
		content=urlopen(req)
	except:
		return False
	
	return True#当所有评教提交后，才返回True标志



#=========================================================
def quit_evaluation(url):
	print("一定时间内，你可打开如下链接确认你的评教情况：")
	print(url)
	input("回车结束。谢谢使用！")
	exit()
	


#=========================================================
if __name__=="__main__":
	print(author_info)

	#登录
	id=input("请输入用户名：")
	key=input("请输入密码：")
	login_info={
			"__VIEWSTATE":"dDwtMTcyNDQ4MTQ0ODs7Ppsymdh1zt1FLm27hnkJ5NY2YVya",
			"TextBox1":id,#username
			"TextBox2":key,#password
			"RadioButtonList1":"%D1%A7%C9%FA",
			"Button1":"",
			"lbLanguage":"",
	}
	
	print("尝试登录服务器……")
	for n in range(100):#鉴于某些学校评教时网络太差,for循环100次登录服务器。
		main_page=login(login_info)
		if main_page!=False:break
	
	if main_page==False:
		print("登录失败。请检查用户名密码后重试。")
		print("有时因网络和服务器的原因而一直无法登录，稍后重试即可。")
		exit()
	else:
		print("成功。")
	

	#确认需要评教的课程
	courses=get_course_list(main_page[1])#get courses list
	if courses=={}:
		print("未搜索到任何评教课程。")
		print("评教未开始或者你已完成评教。")
		#quit_evaluation(main_page[0])

	print("")
	print("以下所列为搜索到的课程：")
	print("========================")
	for i in courses:
		print(courses[i])
	print("")
	input("请核对是否正确。（回车表示确认，ctrl+c终止退出）")


	#评教
	print("正在执行……")
	for i in courses:
		print("处理："+courses[i])
		v=set_evaluation(courses[i],i,main_page[0])
		if v==False:
			#失败
			break
		elif v==True:
			print("请稍等，核对评教结果……")
			result=login(login_info)
			if get_course_list(result[1])=={}:#再次搜索主页中是否有评教科目了。没有表示评教成功。
				print("所有课程评教完毕！")
				quit_evaluation(result[0])
			else:#可能评教失败了
				print("未知状况。")
				quit_evaluation(result[0])
	#失败
	print("评教失败。")
	quit_evaluation(main_page[0])
