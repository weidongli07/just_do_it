# -*- coding: utf-8 -*-
"""
Created on Sat May  7 10:18:06 2022

@author: PC
"""

import pandas as pd
import pymysql 
import tkinter as tk
from tkinter import filedialog
import os
import sys
import urllib.request
from pathlib import Path


#def getLocalFile():#通过浏览文件获取文件路径
#    root=tk.Tk()
#    root.withdraw()
#
#    filePath=filedialog.askopenfilename()
#
#    print('文件路径：',filePath)
#    return filePath
#
#filepath = getLocalFile()
#father_path = os.path.abspath(os.path.dirname(filepath)+os.path.sep+".")
#filename = filepath.split("/")[-1].split(".")[0]

conn = pymysql.connect(
    host='159.75.235.206',
    port=7166,
    user='inq_liweidong',
    password='hCSFxNBb4WN0dqt0eUhZ',
    charset='utf8')
print('数据库连接成功')

filepath = 1
# phone_list = tuple(input('输入手机号列表:').split(','))
# start_date = input('订单开始日期：')
# end_date = input('订单结束日期：')
start_date = '2022-05-15'
end_date = '2022-06-01'
phone_list = 1
#tuple(pd.rqead_excel(filepath).sender_phone)

sql = '''SELECT
	o.forder_id '订单ID',
	o.forder_num '订单编号',
    o.fsender_phone '物流手机号',
	DATE_FORMAT( o.forder_time, '%Y-%m-%d %H:%i:%s' ) AS '下单时间',
	g.flast_quote / 100 "下单金额",
	st.forder_status_name "下单状态",
	o.Fproduct_name "机型",
	cla.fname "类目",
	ca.Fcategory_name "品牌",
	DATE_FORMAT( o.Fpay_out_time, '%Y-%m-%d %H:%i:%s' ) AS "付款时间",
	o.Fpay_out_price / 100 "付款金额",
	(
CASE
	
	WHEN o.fsupply_partner = 1 THEN
	'闪修侠' 
	WHEN o.fsupply_partner = 2 THEN
	'小站（自营）' 
	WHEN o.fsupply_partner = 3 THEN
	'小站（加盟）' 
	WHEN o.fsupply_partner = 4 THEN
	'速回收' 
	WHEN o.fsupply_partner = 5 THEN
	'小豹哥' 
	WHEN o.fsupply_partner = 6 THEN
	'顺丰' 
	WHEN ( o.fsupply_partner = 0 ) 
	AND Fbusiness_mode = 13 THEN
		'服务商回收' -- 华为商城渠道的服务商回收,包括电视和非笔记本、手机、平板类目
		
		WHEN ( o.fsupply_partner = 0 ) 
		AND o.Frecycle_type = 1 THEN
			'顺丰邮寄' -- 邮寄没有履约方
--  	 when fsupply_partner= 0 and Frecycle_type = 2 and Forder_status = 80 then '上门单已取消' -- 没有履约方
--  	 when fsupply_partner= 0 and Frecycle_type = 3 and Forder_status = 80 then '到店单已取消' -- 没有履约方
-- 	else '其他' -- 履约方为空的多为未完成的订单
			
		END 
		) AS '履约方',
		(
		CASE
				
				WHEN Fbusiness_mode = 0 THEN
				'0' 
				WHEN Fbusiness_mode = 1 THEN
				'B端帮卖' 
				WHEN Fbusiness_mode = 2 THEN
				'C端帮卖' 
				WHEN Fbusiness_mode = 3 THEN
				'C端回收' 
				WHEN Fbusiness_mode = 4 THEN
				'B端回收' 
				WHEN Fbusiness_mode = 5 THEN
				'以旧换新' 
				WHEN Fbusiness_mode = 6 THEN
				'捐赠' 
				WHEN Fbusiness_mode = 7 THEN
				'B端帮卖（基础）' 
				WHEN Fbusiness_mode = 8 THEN
				'C端帮卖（基础）' 
				WHEN Fbusiness_mode = 9 THEN
				'以旧换新（基础）' 
				WHEN Fbusiness_mode = 10 THEN
				'保值回购' 
				WHEN Fbusiness_mode = 11 THEN
				'保值回购转回收单' 
				WHEN Fbusiness_mode = 12 THEN
				'售后订单' 
				WHEN Fbusiness_mode = 13 THEN
				'服务商回收订单' 
				WHEN Fbusiness_mode = 14 THEN
				'B2B2C售后订单' 
				WHEN Fbusiness_mode = 15 THEN
				'荣耀保值换新' 
				WHEN Fbusiness_mode = 16 THEN
				'上门换新' 
				WHEN Fbusiness_mode = 17 THEN
				'竞拍' 
			END 
			) AS '业务类型',
			ste.Fstore_name '门店名称',
		CASE
				o.Frecycle_type 
				WHEN 1 THEN
				'邮寄' 
				WHEN 2 THEN
				'上门' 
				WHEN 3 THEN
				'到店' 
				WHEN 4 THEN
				'ATM' 
			END '回收类型',
	cd.fchannel_name,
	cd.ftag_name 
FROM
	recycle.t_order o
	INNER JOIN (
	SELECT
		tm.fp_id AS fpid,
		tc.fchannel_id,
		tc.fchannel_name,
		tt.ftag_id,
		tt.ftag_name 
	FROM
		recycle.t_channel AS tc
		LEFT JOIN recycle.t_tag AS tt ON tt.fchannel_id = tc.fchannel_id
		LEFT JOIN recycle.t_maptag AS tm ON tm.ftag_id = tt.ftag_id 
	WHERE
		tc.fchannel_id IN (
			40000001,
			30000001,
			10000011,
			10000012,
			10000001,
			10000039,
			10000040,
			10000056,
			10000057,
			10000058,
			10000060,
			10000080,
			10000082,
			10000083,
			10000118,
			10000120,
			10000128,
			10000147,
			10000211,
			10000187,
			10000197,
			10000079,
			10000177,
			10000246,
			10000257,
			10000333,
			10000306,
			10000016,
			10000021,
			10000035,
			10000036,
			10000063,
			10000099,
			10000100,
			10000102,
			10000113,
			10000121,
			10000130,
			10000137,
			10000144,
			10000171,
			10000174,
			10000212,
			10000245,
			10000248,
			10000249,
			10000250,
			10000201,
			10000265,
			10000302,
			10000303,
			10000332,
			10000334,
			10000336,
			10000337,
			10000134,
			10000311,
			10000313,
			10000350,
			10000352,
			10000353,
			10000354,
			10000355,
			10000356,
			10000357,
			10000359,
			10000360,
			10000361,
			10000362,
			10000363,
			10000364,
			10000358,
			10000366,
			10000374,
			10000375,
			10000377,
			10000378,
			10000381,
			10000382,
			10000383,
			10000384,
			10000385,
			10000385,
			10000388,
			10000387,
			10000389,
			10000390,
			10000394,
			10000368,
			10000369,
			10000370,
			10000371,
			10000372,
			10000373,
			10000391,
			10000376,
			10000392,
			10000393,
			10000379,
			10000380,
			10000395,
			10000396,
			10000397,
			10000398,
			10000399,
			10000400,
			10000401,
			10000402,
			10000403,
			10000404,
			10000405,
			10000406,
			10000407,
			10000408,
			10000409,
			10000410,
			10000412,
			10000413,
			10000414,
			10000415,
			10000416,
			10000417,
			10000418,
			10000419,
			10000420,
			10000421,
			10000422,
			10000423,
			10000424,
			10000425,
			10000426,
			10000349,
			10000340,
			10000338,
			10000876,
			10000943 
		) 
	) cd ON cd.fpid = o.fpid
	LEFT JOIN recycle.t_order_status st ON o.Forder_status = st.Forder_status_id
	LEFT JOIN recycle.t_product p ON o.Fproduct_id = p.Fproduct_id
	LEFT JOIN recycle.t_category ca ON ca.Fcategory_id = p.Fcategory_id
	LEFT JOIN recycle.t_pdt_class cla ON cla.Fid = p.Fclass_id
	LEFT JOIN recycle.t_order_snapshot sn ON sn.Forder_id = o.Forder_id
	LEFT JOIN hjxmba_db.t_store_info ste ON ste.Fstore_id = sn.Fo1_id
	LEFT JOIN ( SELECT g.Fvaluation, g.Flast_quote, g.Fgoods_id FROM recycle.t_goods g ) g ON o.Fgoods_id = g.Fgoods_id 
WHERE
	DATE_FORMAT( o.forder_time, '%Y-%m-%d' ) BETWEEN '{1}' 
AND '{2}'
-- and o.fsender_phone in {0}
and cd.fchannel_name = 'vivo商城'
-- limit 100 '''.format(phone_list,start_date,end_date)
df = pd.read_sql(sql, conn)
#df.to_excel(father_path +'\\'+ filename + '_result'+ '.xlsx')
df.to_excel(r"D:\Work\回收宝\SQL\python\vivo商城订单明细_20220515.xlsx")
