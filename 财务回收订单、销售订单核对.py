

import numpy as np
import pymysql
import tkinter as tk
from tkinter import filedialog
import os
import sys
import pandas as pd
import urllib.request
from pathlib import Path


def getLocalFile():#通过浏览文件获取文件路径
    root=tk.Tk()
    root.withdraw()

    filePath=filedialog.askopenfilename()

    print('文件路径：',filePath)
    return filePath


'''
1.闲鱼小站闲鱼小程序的小站履约  标记为非自有
2. 履约方为速回收，小豹哥，顺丰上门履约的标记为自有上门
3. 标记是否为测试单
'''

conn = pymysql.connect(
host='159.75.235.206',
port=7166,
user='inq_liweidong',
password='hCSFxNBb4WN0dqt0eUhZ',
charset='utf8')

conn_xz = pymysql.connect(
host='159.75.235.206',
port=7058,
user='inq_liweidong',
password='hCSFxNBb4WN0dqt0eUhZ',
charset='utf8')

print('数据库连接成功')

###--------------------------------------------------------------回收------------------------------------------------------

filepath_hs = getLocalFile()
father_path_hs = os.path.abspath(os.path.dirname(filepath_hs)+os.path.sep+".")
filename_hs = filepath_hs.split("/")[-1].split(".")[0]
df = pd.read_excel(filepath_hs,sheet_name='Sheet2')
orderlists = tuple(df.loc[:,'clean订单号'].values.tolist())

sql = '''SELECT
 	o.Forder_num,
 	(case when fsupply_partner = 2 then '小站(自营)' when  fsupply_partner = 3 then '小站(加盟)' else fsupply_partner end) as '门店类型',
 	(case when fsupply_partner = 1 then "闪修侠" when fsupply_partner in (2,3) then "小站" when fsupply_partner = 4 then "速回收" when fsupply_partner = 5 then "小豹哥" when fsupply_partner = 6 then "顺丰"
 	else "邮寄" end) as 履约方,
   	case o.Frecycle_type
        when 1 then '邮寄'
        when 2 then '上门'
        when 3 then '到店'
        when 4 then 'ATM'
        end '回收类型',
 	o.Forder_id "订单ID",
 	o.Fseries_number "条码",
 	c.Fchannel_name "渠道",
 	(
 	CASE
 			
 			WHEN c.Fchannel_name IN (
				'回收宝主站',
				'H5回收宝',
				'2C百度竞价',
				'2C今日头条',
				'2C神马竞价',
				'官方微博',
				'360',
				'搜狗',
				'2C百度网盟',
				'2C广点通1',
				'2C广点通2',
				'资讯站 zx.huishoubao.com',
				'辣子机公众号',
				'竞价平台',
				'付费营销渠道',
				'可乐优品商城',
				'手机百度合作渠道',
				'自有回收员工内购' 
				) THEN
				'PC&H5' 
				WHEN c.Fchannel_name IN ( '小豹帮卖') then '小豹帮卖'
				WHEN c.Fchannel_name IN ( '回收宝APP-2C', '自有-拍立卖', '2C回收宝APP-iOS', '自有回收-抖音短视频' ) THEN
				'APP' 
				WHEN c.Fchannel_name IN ( '微信公众号2', '微信公众号活动', '小豹福利社微信小程序', '自有-清库存专用渠道' ) THEN
				'微信小程序' 
				WHEN c.Fchannel_name IN (
 					'安兔兔',
 					'2C中兴努比亚',
 					'2C 爱锋派',
 					'2C iTools',
 					'2CiPhone频道',
 					'2C 飞蚂蚁',
 					'2C 鲁大师',
 					'2C刷机精灵',
 					'闪修侠',
 					'百度手机助手',
 					'魅族生活App',
 					'蓝店',
 					'即有分期',
 					'2C 爱奇艺',
 					'中国移动手机商城-广东',
 					'中国移动-积分商城',
 					'2C 爱奇艺积分商城',
 					'快应用1',
 					'2C 92回收',
 					'闲鱼小站本地合作',
 					'vivo保值回收',
 					'vivo一键换机',
 					'哎咆科技',
 					'全民钱包',
 					'移动CPS渠道',
 					'北京114',
 					'闲鱼小站闲鱼小程序',
 					'小站加盟（占位渠道4）' ,
 					'2C小黄狗APP',
                    '一一米淘',
                    '闪修侠用户端',
                    '速回收-工程师挖单',
						'小站加盟（占位渠道6）',
						'房张辉（郑州虚拟商户）',
						'王成（珠海中山虚拟门店）',
						'田昱杭（重庆虚拟商户）',
						'耿永坤（重庆虚拟商户）',
						'高明军（广州虚拟商户）',
						'陈四化（广州虚拟商户）',
						'曾冬冬（苏州虚拟商户）',
						'曾建国（南京虚拟商户）',
						'韩田丰（惠州虚拟商户）',
						'岳新闻（北京虚拟商户）',
						'孙景次（深圳虚拟商户）',
						'刘珊珊（哈尔滨虚拟商户）',
						'张小标（上海虚拟商户）',
						'陈永飞（佛山虚拟商户）',
						'曹昆昆（徐州虚拟商户） ',
						'王彬（成都虚拟门店）',
						'曾超全（东莞虚拟门店）',
						'曾秀丽（佛山虚拟门店）',
						'王华（福州虚拟门店）',
						'陈家乐（贵阳虚拟门店）',
						'曾洋（杭州虚拟门店）',
						'李珍（ 南京虚拟门店）',
						'张大伟（南宁虚拟门店）',
						'张晓静（泉州虚拟门店）',
						'陈涛（厦门虚拟门店）',
						'李永见（苏州虚拟门店）',
						'曾陈杰（上海虚拟门店）',
						'王长海（天津虚拟门店）',
						'苏威（温州虚拟门店）',
						'史伟强（烟台虚拟商户）',
						'余勇（驻马店虚拟商户）',
						'余小威（北京虚拟商户）',
						'吴凯（北京虚拟商户）',
						'杜国伟（北京虚拟商户）',
						'张莉莉（北京虚拟商户）',
						'余小佩（北京虚拟商户）',
						'邹亮（成都虚拟商户）',
						'张治强（大连虚拟商户）',
						'王袆（大庆虚拟商户）',
						'于荣圣（德州虚拟商户）',
						'刘慧（福州虚拟商户）',
						'陈四化（广州虚拟商户）',
						'燕永鹏（佛山虚拟商户）',
						'曾亚伟（杭州虚拟商户）',
						'蒋晓妍（合肥虚拟商户）',
						'曾飞翔（惠州虚拟商户）',
						'郝孝祯（济南虚拟商户）',
						'韩田丰（廊坊虚拟商户',
						'孙志明（辽阳虚拟商户）',
						'贺拥军（洛阳虚拟商户）',
						'蒋静静（昆明虚拟商户）',
						'苏亚伟（南昌虚拟商户）',
						'李飞（南昌虚拟商户）',
						'岳新伟（宁波虚拟商户）',
						'杨张龙（青岛虚拟商户）',
						'曾召春（上海虚拟商户）',
						'王稳稳（上海虚拟商户）',
						'燕明禄（上海虚拟商户）',
						'曾可（深圳虚拟商户）',
						'孙吉衡（深圳虚拟商户）',
						'陈柳（深圳虚拟商户）',
						'田野（沈阳虚拟商户）',
						'赵金源（太原虚拟商户）',
						'陈阳（无锡虚拟商户）',
						'曾倩倩（无锡虚拟商户）',
						'申志见（武汉虚拟商户）',
						'李树广（西安虚拟商户）',
						'万海芸（新余虚拟商户）',
						'卢永山（义乌金华虚拟商户）',
						'刘华玉（长春虚拟商户）',
						'张东伟（长沙虚拟商户）',
						'曹诗普（镇江虚拟商户）' 
 					) THEN
 					'CPS中小渠道' 
 					WHEN c.Fchannel_name IN ( "vivo商城" ) THEN
 					"vivo商城" 
 					WHEN c.Fchannel_name IN ( "分期乐" ) THEN
 					"分期乐"  
 					WHEN c.Fchannel_name IN ( "华为商城回收" ) THEN
 					"华为商城"
                      WHEN c.Fchannel_name IN ( "荣耀商城" ) THEN
 					"荣耀商城"
                      WHEN c.Fchannel_name IN ( "官网H5" ) THEN
 					  "官网H5"
                      WHEN c.fchannel_id IN ( 10000056, 10001054 ) THEN
                        '官方微博'
                        WHEN c.fchannel_id IN ( 10001040 ) THEN
                        '联想商城'
                        WHEN c.fchannel_id IN ( 10001070 ) THEN
                        '联想管家'
                        WHEN c.fchannel_id IN ( 10000170 ) THEN
                        '支付宝小程序'
 					ELSE "非自有"
 					END 
 					) AS '归属渠道',
 					( CASE WHEN o.ftest = 1 THEN "测试订单" ELSE "正常订单" END ) AS "是否为测试单",
 					CAST(xyd.Fxy_order_id AS CHAR) "闲鱼订单ID",
 					date_format(o.Fpay_out_time,'%Y-%m-%d' ) "付款时间",
 					o.Fpay_out_price / 100 "付款金额",
 					(case when (ac.Faccount_id is not null or o.Fchannel_id in (10000257)) then "滞留单" else "否" end) AS "是否为滞留单" 
						FROM
 							recycle.t_order o
 							LEFT JOIN recycle.t_order_snapshot sh ON sh.forder_id = o.forder_id
 							LEFT JOIN recycle.t_xy_order_data xyd ON o.forder_id = xyd.forder_id
 							LEFT JOIN recycle.t_channel c ON c.Fchannel_id = o.Fchannel_id
   							LEFT JOIN (select Faccount_id from  recycle.t_account_info where Faccount = 'wendylei@huishoubao.com.cn') ac ON ac.Faccount_id = o.Faccount_id
						WHERE
  o.forder_num in {0}'''.format(orderlists)


sql_2 = '''select
xy.Forder_id
,xy.Forder_num
,xy.Fstore_id
,s.Fstore_name
,xy.Fclerk_name
from recycle.t_xyxz_order xy
left join hjxmba_db.t_store_info s on s.Fstore_id = xy.Fstore_id
where  xy.Forder_num in {0}'''.format(orderlists)

print('开始查询sql')
stime = datetime.now()

df_sql = pd.read_sql(sql,conn)

df_sql_2 = pd.read_sql(sql_2,conn_xz)


etime = datetime.now()
print('结束查询，总查询用时：',(etime - stime).seconds,'s')

df_detail_hs = pd.merge(left=df_sql, right=df_sql_2[['Forder_id','Fstore_name','Fclerk_name']],how='left',left_on='订单ID',right_on='Forder_id')
df_detail_hs.to_excel(father_path_hs +'\\'+ filename_hs + '_result'+ '.xlsx')


##-------------------------------------------------------------------销售---------------------------------------------


# filepath_xs = getLocalFile()
# father_path_xs = os.path.abspath(os.path.dirname(filepath_xs)+os.path.sep+".")
# filename_xs = filepath_xs.split("/")[-1].split(".")[0]
# df_sr = pd.read_excel(filepath_xs,sheet_name='Sheet2')
# tmlists = tuple(df_sr.loc[:,'CLEAN原条码'].values.tolist())

# sql_1 = '''SELECT
#  	o.Forder_num,
#  	 	(case when fsupply_partner = 2 then '小站(自营)' when  fsupply_partner = 3 then '小站(加盟)' else fsupply_partner end) as '门店类型',
#  	(case when fsupply_partner = 1 then "闪修侠" when fsupply_partner in (2,3) then "小站" when fsupply_partner = 4 then "速回收" when fsupply_partner = 5 then "小豹哥" when fsupply_partner = 6 then "顺丰"
#  	else "邮寄" end) as 履约方,
#          	case o.Frecycle_type
# when 1 then '邮寄'
# when 2 then '上门'
# when 3 then '到店'
# when 4 then 'ATM'
# end '回收类型',
#  	o.Forder_id "订单ID",
#  	o.Fseries_number "条形码",
#  	c.Fchannel_name "渠道",
#  	(
#  	CASE
 			
#  			WHEN c.Fchannel_name IN (
# 				'回收宝主站',
# 				'H5回收宝',
# 				'2C百度竞价',
# 				'2C今日头条',
# 				'2C神马竞价',
# 				'官方微博',
# 				'360',
# 				'搜狗',
# 				'2C百度网盟',
# 				'2C广点通1',
# 				'2C广点通2',
# 				'资讯站 zx.huishoubao.com',
# 				'辣子机公众号',
# 				'竞价平台',
# 				'付费营销渠道',
# 				'可乐优品商城',
# 				'手机百度合作渠道',
# 				'自有回收员工内购' 
# 				) THEN
# 				'PC&H5' 
# 				WHEN c.Fchannel_name IN ( '小豹帮卖') then '小豹帮卖'
# 				WHEN c.Fchannel_name IN ( '回收宝APP-2C', '自有-拍立卖', '2C回收宝APP-iOS', '自有回收-抖音短视频' ) THEN
# 				'APP' 
# 				WHEN c.Fchannel_name IN ( '微信公众号2', '微信公众号活动', '小豹福利社微信小程序', '自有-清库存专用渠道' ) THEN
# 				'微信小程序' 
# 				WHEN c.Fchannel_name IN (
#  					'安兔兔',
#  					'2C中兴努比亚',
#  					'2C 爱锋派',
#  					'2C iTools',
#  					'2CiPhone频道',
#  					'2C 飞蚂蚁',
#  					'2C 鲁大师',
#  					'2C刷机精灵',
#  					'闪修侠',
#  					'百度手机助手',
#  					'魅族生活App',
#  					'蓝店',
#  					'即有分期',
#  					'2C 爱奇艺',
#  					'中国移动手机商城-广东',
#  					'中国移动-积分商城',
#  					'2C 爱奇艺积分商城',
#  					'快应用1',
#  					'2C 92回收',
#  					'闲鱼小站本地合作',
#  					'vivo保值回收',
#  					'vivo一键换机',
#  					'哎咆科技',
#  					'全民钱包',
#  					'移动CPS渠道',
#  					'北京114',
#  					'闲鱼小站闲鱼小程序',
#  					'小站加盟（占位渠道4）' ,
#  					'2C小黄狗APP',
#                     '一一米淘',
#                     '闪修侠用户端',
#                     '速回收-工程师挖单',
# 						'小站加盟（占位渠道6）',
# 						'房张辉（郑州虚拟商户）',
# 						'王成（珠海中山虚拟门店）',
# 						'田昱杭（重庆虚拟商户）',
# 						'耿永坤（重庆虚拟商户）',
# 						'高明军（广州虚拟商户）',
# 						'陈四化（广州虚拟商户）',
# 						'曾冬冬（苏州虚拟商户）',
# 						'曾建国（南京虚拟商户）',
# 						'韩田丰（惠州虚拟商户）',
# 						'岳新闻（北京虚拟商户）',
# 						'孙景次（深圳虚拟商户）',
# 						'刘珊珊（哈尔滨虚拟商户）',
# 						'张小标（上海虚拟商户）',
# 						'陈永飞（佛山虚拟商户）',
# 						'曹昆昆（徐州虚拟商户） ',
# 						'王彬（成都虚拟门店）',
# 						'曾超全（东莞虚拟门店）',
# 						'曾秀丽（佛山虚拟门店）',
# 						'王华（福州虚拟门店）',
# 						'陈家乐（贵阳虚拟门店）',
# 						'曾洋（杭州虚拟门店）',
# 						'李珍（ 南京虚拟门店）',
# 						'张大伟（南宁虚拟门店）',
# 						'张晓静（泉州虚拟门店）',
# 						'陈涛（厦门虚拟门店）',
# 						'李永见（苏州虚拟门店）',
# 						'曾陈杰（上海虚拟门店）',
# 						'王长海（天津虚拟门店）',
# 						'苏威（温州虚拟门店）',
# 						'史伟强（烟台虚拟商户）',
# 						'余勇（驻马店虚拟商户）',
# 						'余小威（北京虚拟商户）',
# 						'吴凯（北京虚拟商户）',
# 						'杜国伟（北京虚拟商户）',
# 						'张莉莉（北京虚拟商户）',
# 						'余小佩（北京虚拟商户）',
# 						'邹亮（成都虚拟商户）',
# 						'张治强（大连虚拟商户）',
# 						'王袆（大庆虚拟商户）',
# 						'于荣圣（德州虚拟商户）',
# 						'刘慧（福州虚拟商户）',
# 						'陈四化（广州虚拟商户）',
# 						'燕永鹏（佛山虚拟商户）',
# 						'曾亚伟（杭州虚拟商户）',
# 						'蒋晓妍（合肥虚拟商户）',
# 						'曾飞翔（惠州虚拟商户）',
# 						'郝孝祯（济南虚拟商户）',
# 						'韩田丰（廊坊虚拟商户',
# 						'孙志明（辽阳虚拟商户）',
# 						'贺拥军（洛阳虚拟商户）',
# 						'蒋静静（昆明虚拟商户）',
# 						'苏亚伟（南昌虚拟商户）',
# 						'李飞（南昌虚拟商户）',
# 						'岳新伟（宁波虚拟商户）',
# 						'杨张龙（青岛虚拟商户）',
# 						'曾召春（上海虚拟商户）',
# 						'王稳稳（上海虚拟商户）',
# 						'燕明禄（上海虚拟商户）',
# 						'曾可（深圳虚拟商户）',
# 						'孙吉衡（深圳虚拟商户）',
# 						'陈柳（深圳虚拟商户）',
# 						'田野（沈阳虚拟商户）',
# 						'赵金源（太原虚拟商户）',
# 						'陈阳（无锡虚拟商户）',
# 						'曾倩倩（无锡虚拟商户）',
# 						'申志见（武汉虚拟商户）',
# 						'李树广（西安虚拟商户）',
# 						'万海芸（新余虚拟商户）',
# 						'卢永山（义乌金华虚拟商户）',
# 						'刘华玉（长春虚拟商户）',
# 						'张东伟（长沙虚拟商户）',
# 						'曹诗普（镇江虚拟商户）' 
#  					) THEN
#  					'CPS中小渠道' 
#  					WHEN c.Fchannel_name IN ( "vivo商城" ) THEN
#  					"vivo商城" 
#  					WHEN c.Fchannel_name IN ( "分期乐" ) THEN
#  					"分期乐"  
#   					 WHEN c.Fchannel_name IN ( "华为商城回收" ) THEN
#   					 "华为商城"
#                         WHEN c.Fchannel_name IN ( "荣耀商城" ) THEN
#  					  "荣耀商城"
#                       WHEN c.Fchannel_name IN ( "官网H5" ) THEN
#  					  "官网H5"
#                                 WHEN c.fchannel_id IN ( 10000056, 10001054 ) THEN
#         '官方微博'
#         WHEN c.fchannel_id IN ( 10001040 ) THEN
#         '联想商城'
#         WHEN c.fchannel_id IN ( 10001070 ) THEN
#         '联想管家'
#         WHEN c.fchannel_id IN ( 10000170 ) THEN
#         '支付宝小程序'
#   					 ELSE "非自有"
#   					 END 
#  					) AS '归属渠道',
#       o.Fpid,
#  					( CASE WHEN o.ftest = 1 THEN "测试订单" ELSE "正常订单" END ) AS "是否为测试单",
#                     CAST(xyd.Fxy_order_id AS CHAR) "闲鱼订单ID",
#  					date_format(o.Fpay_out_time,'%Y-%m-%d' ) "付款时间",
#  					o.Fpay_out_price / 100 "付款金额",
#  							(case when (ac.Faccount_id is not null or o.Fchannel_id in (10000257)) then "滞留单" else "否" end) AS "是否为滞留单" 
# 						FROM
#  							recycle.t_order o
#  							LEFT JOIN recycle.t_order_snapshot sh ON sh.forder_id = o.forder_id
#                           LEFT JOIN recycle.t_xy_order_data xyd ON o.forder_id = xyd.forder_id

#  							LEFT JOIN recycle.t_channel c ON c.Fchannel_id = o.Fchannel_id
#  							LEFT JOIN (select Faccount_id from  recycle.t_account_info where Faccount = 'wendylei@huishoubao.com.cn') ac ON ac.Faccount_id = o.Faccount_id
# 						WHERE
# o.Fseries_number in {0}'''.format(tmlists)

# print('开始查询sql')
# stime = datetime.now()
# df_sql_1 = pd.read_sql(sql_1,conn)

# sql_3 = '''select
# xy.Forder_id
# ,xy.Forder_num
# ,xy.Fserial_number
# ,xy.Fstore_id
# ,s.Fstore_name
# ,xy.Fclerk_name
# from recycle.t_xyxz_order xy
# left join hjxmba_db.t_store_info s on s.Fstore_id = xy.Fstore_id
# where  xy.Forder_id in {0}'''.format(tuple(df_sql_1['订单ID'].values.tolist()))
# df_sql_3 = pd.read_sql(sql_3,conn_xz)


# etime = datetime.now()
# print('结束查询，总查询用时：',(etime - stime).seconds,'s')

# df_detail_xs = pd.merge(left=df_sql_1, right=df_sql_3[['Forder_id','Fstore_name','Fclerk_name']],how='left',left_on='订单ID',right_on='Forder_id')
# df_detail_xs.to_excel(father_path_xs +'\\'+ filename_xs + '_result'+ '.xlsx')







