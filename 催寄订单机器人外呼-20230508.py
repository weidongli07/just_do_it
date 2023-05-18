# -*- coding: utf-8 -*-
"""
Created on Mon Aug  8 18:15:07 2022

@author: PC
"""
import pandas as pd
import numpy as np
import pymysql 
from datetime import datetime,timedelta
import tkinter as tk
from tkinter import filedialog
import os
import sys
import json
import requests
import prettytable as pt
from apscheduler.schedulers.blocking import BlockingScheduler
import Sobotcall

def main():

    
    today = datetime.today().strftime('%m月%d日')
    yesterday = (datetime.now() + timedelta(days=-1)).strftime("%Y年%m月%d日")
    the_2day_before = (datetime.now() + timedelta(days=-2)).strftime("%Y年%m月%d日")
    the_3day_before = (datetime.now() + timedelta(days=-3)).strftime("%Y年%m月%d日")
    
    
    month_frist_day = (datetime.now() + timedelta(days=-1)).replace(day=1).strftime("%Y-%m-%d")
    the_7day_before = (datetime.now() + timedelta(days=-7)).strftime("%Y-%m-%d")
    
    begin_time = '2022-08-01'
    end_time = datetime.now().strftime("%Y-%m-%d")
    
    
    # def getLocalFile():#通过浏览文件获取文件路径
    #     root=tk.Tk()
    #     root.withdraw()
    
    #     filePath=filedialog.askopenfilename()
    
    #     print('已选择文件路径：',filePath)
    #     return filePath
    
    
    # record_file_path = getLocalFile()
    record_file_path = r"E:\自有后端\自付款失败记录.xlsx"
    father_record_file_path = os.path.abspath(os.path.dirname(record_file_path)+os.path.sep+".")
    
    
    # record_file_path = r'E:\自有后端\自付款失败记录.xlsx'
    channel_zy =( 40000001, 30000001, 10000011, 10000012, 10000001, 10000039, 10000040, 10000056, 10000057, 10000058, 10000060, 10000080, 10000082, 10000083, 10000118, 10000120, 10000128, 10000147, 10000211, 10000187, 10000197, 10000079, 10000177, 10000246, 10000257, 10000333, 10000306, 10000016, 10000021, 10000035, 10000036, 10000063, 10000099, 10000100, 10000102, 10000113, 10000121, 10000130, 10000137, 10000144, 10000171, 10000174, 10000212, 10000245, 10000248, 10000249, 10000250, 10000201, 10000265, 10000302, 10000303, 10000332, 10000334, 10000336, 10000337, 10000134, 10000311, 10000313, 10000350, 10000352, 10000353, 10000354, 10000355, 10000356, 10000357, 10000359, 10000360, 10000361, 10000362, 10000363, 10000364, 10000358, 10000366, 10000374, 10000375, 10000377, 10000378, 10000381, 10000382, 10000383, 10000384, 10000385, 10000385, 10000388, 10000387, 10000389, 10000390, 10000394, 10000368, 10000369, 10000370, 10000371, 10000372, 10000373, 10000391, 10000376, 10000392, 10000393, 10000379, 10000380, 10000395, 10000396, 10000397, 10000398, 10000399, 10000400, 10000401, 10000402, 10000403, 10000404, 10000405, 10000406, 10000407, 10000408, 10000409, 10000410, 10000412, 10000413, 10000414, 10000415, 10000416, 10000417, 10000418, 10000419, 10000420, 10000421, 10000422, 10000423, 10000424, 10000425, 10000426, 10000349, 10000340, 10000338, 10000876, 10000943,10001039,10001041,10001040,10001054,10000170,10001070)
    black_phone_list = [18575677997, 18665856610, 13342951610, 17688938632, 15538079080, 18123771160, 15296003228, 15779251532, 15918718989, 13760148884, 13420133645, 13129616809, 18296552227, 13723739215, 13377677107, 15986637346, 13057670209, 15999579786, 15820756400, 18583702158, 18826406044, 15201903213, 13918257799,19327452289]
    
    conn = pymysql.connect(
        
    host='159.75.235.206',
    port=7166,
    user='inq_liweidong',
    password='hCSFxNBb4WN0dqt0eUhZ',
    charset='utf8')
    
    
    # 订单明细sql
    sql_order = """ 
    SELECT
    	o.forder_id,
    	o.forder_num,
    	o.Fsender_phone,
        u.Fphone_num  '用户注册手机号',
    o.forder_time AS '下单时间',
    ccl.Fupdate_time AS '取消时间',
    		
    lr.Fvisit_time '预约上门时间',
    lr.Fdeliver_name '寄件用户名',
    o.Fsend_time '寄出时间',
    o.Fgetin_time '收货时间',
    o.Fbargin_time '议价时间',
    -- 		ccl.Fremark AS '取消原因',
    		r.g_Freason AS '取消原因',
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
    			-- ste.Fstore_name '门店名称',
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
    
    	cd.ftag_name ,
        cd.fchannel_id,
        	cd.fchannel_name,
            cd.channel,
            case 
            when cd.channel in ('H5','PC','可乐优品商城','APP_android','APP_ios','微信小程序','估价未下单挽回活动') then '自有'
                    when cd.channel in ('投放H5-抖音','投放H5-搜索引擎') then '投放H5'
                    when cd.channel in ('CPS中小渠道','分期乐') then 'CPS中小渠道'
                    else cd.channel end as 'main_channel'
    FROM
    	recycle.t_order o
    	INNER JOIN (
    	SELECT
    		tm.fp_id AS fpid,
    		tc.fchannel_id,
    		tc.fchannel_name,
    					        (case
            when  tm.fp_id = 1588 then '估价未下单挽回活动'
            when tc.fchannel_id = 10000943 and tm.fp_id in (11030,11031,11032,11033,11034,11035) then '投放H5-抖音'
            when tc.fchannel_id = 10000943  then '投放H5-搜索引擎'
            when (tm.fp_id in (1272,1380,1180,1273,1053,1368,1182,1379,1356,1177,1105,1331,1104) or tc.fchannel_id in (30000001,10000246)) then 'H5'
            when (tm.fp_id in (1042,1367,1146,1330,1181,1147) or tc.fchannel_id in (40000001)) then 'PC'
            when tm.fp_id in (1588,1587) or tc.fchannel_id in (10000177) then '可乐优品商城'
            when tc.fchannel_id in (10000060) and tm.fp_id not in (1260, 1176) then 'APP_android'
            when tm.fp_id in (1260, 1176) then 'APP_ios'
            when tc.fchannel_id in (10000012, 10000001) then '微信小程序'
            when tc.fchannel_id in (10000113,	10000212,	10000102,	10000265,	10000016,	10000134,	10000174,	10000021,	10000036,	10000334,	10000144,	10000171,	10000201,	10000303,	10000130,	10000100,	10000340,	10000099,	10000250,	10000063,10000313,10000311) then 'CPS中小渠道'
            when tc.fchannel_id in (10000337) then '华为'
            when tc.fchannel_id in (10000121) then 'vivo'
            when tc.fchannel_id in (10000137) then '分期乐'
            when tc.fchannel_id in (10000876) then '荣耀'
            when tc.fchannel_id in (10000056,10000054) then '官方微博'
            when tc.fchannel_id in (10001040) then '联想商城'
            else tc.fchannel_name end) as channel,
    		tt.ftag_id,
    		tt.ftag_name 
    	FROM
    		recycle.t_channel AS tc
    		LEFT JOIN recycle.t_tag AS tt ON tt.fchannel_id = tc.fchannel_id
    		LEFT JOIN recycle.t_maptag AS tm ON tm.ftag_id = tt.ftag_id 
    	WHERE
    		tc.fchannel_id IN {2}
    	) cd ON cd.fpid = o.fpid
    	LEFT JOIN recycle.t_order_status st ON o.Forder_status = st.Forder_status_id
    	LEFT JOIN recycle.t_product p ON o.Fproduct_id = p.Fproduct_id
    	LEFT JOIN recycle.t_category ca ON ca.Fcategory_id = p.Fcategory_id
    	LEFT JOIN recycle.t_pdt_class cla ON cla.Fid = p.Fclass_id
    	LEFT JOIN recycle.t_order_snapshot sn ON sn.Forder_id = o.Forder_id
        left join (select  fcreate_time,u.Fuser_id ,u.Funion_id ,u.Fphone_num,u.fopen_id,Fuser_name from recycle.t_user u )u
    	ON u.Fuser_id = o.Fuser_id
    	-- LEFT JOIN hjxmba_db.t_store_info ste ON ste.Fstore_id = sn.Fo1_id
    	LEFT JOIN ( SELECT g.Fvaluation, g.Flast_quote, g.Fgoods_id FROM recycle.t_goods g where DATE_FORMAT(g.Fcreate_time, '%Y-%m-%d' ) between '{0}' and '{1}' ) g ON o.Fgoods_id = g.Fgoods_id 
    	left join  (select
    Forder_id
    ,GROUP_CONCAT(Freason) g_Freason
    from recycle.t_order_remark_record
    
    GROUP BY Forder_id
    ) r on r.Forder_id = o.Forder_id
    left join (
    select
    t.Forder_id
    ,t.Fupdate_time
    ,t.Fremark
    from recycle.t_order_txn t
    inner join (select
    Forder_id
    ,max(Ftxn_id) m_Ftxn_id
    -- Fupdate_time
    from recycle.t_order_txn 
    where Forder_status = 80 and DATE_FORMAT(Fupdate_time, '%Y-%m-%d' ) between '{0}' and '{1}'
    GROUP BY Forder_id
    ) g on g.m_Ftxn_id = t.Ftxn_id
    ) ccl on ccl.Forder_id = o.Forder_id 
    left join (select Forder_id,max(Fvisit_time) as Fvisit_time,max(Fdeliver_name) as Fdeliver_name from  recycle.t_tms_logistics_recycle  where DATE_FORMAT(Forder_time, '%Y-%m-%d' )   between '{0}' and '{1}'  group by Forder_id ) lr on lr.Forder_id = o.Forder_id
    
    WHERE
    DATE_FORMAT( o.forder_time, '%Y-%m-%d' ) between '{0}' and '{1}'
    and o.Fvalid > 0  and o.Ftest = 0
    
    """.format(the_7day_before,end_time,channel_zy)
    
    
    stime = datetime.now()
    print('开始查询sql')
    order_all = pd.read_sql(sql_order,conn)
    etime = datetime.now()
    print(etime,' 结束查询，总查询用时：',(etime - stime).seconds,'s')
    
    # detail.to_excel(r'E:\自有后端\明细.xlsx')
    # detail = order_all.copy()
    detail = order_all.sort_values(by='下单时间').drop_duplicates(subset=['Fsender_phone'],keep='last')
    
    
    def date(para): # 该函数解决读取excel日期格式是数字的问题
        if type(para)==int:
            delta = pd.Timedelta(str(int(para))+'days')
            time = (pd.to_datetime('1899-12-30')+delta).strftime('%m月%d日')
            return time
        else:
            return para
    
    # ------------------------------------------------------------------------
    
    booking_status_list = ['快递柜邮寄','顺丰上门','已下单']
    business_mode = ['C端回收','联想以旧换新']
    recycle_type_list = ['邮寄']
    
    # 根据条件过滤得 预约上门时间过长的回收订单
    order_booking1 = detail[(detail.loc[:,'下单时间'].apply(lambda x:x.strftime('%Y年%m月%d日')) == yesterday)  
                           & (detail.loc[:,'下单状态'].isin(booking_status_list))
                           & (detail.loc[:,'回收类型'].isin(recycle_type_list))
                           & (detail.loc[:,'业务类型'].isin(business_mode))
                           & (detail.loc[:,'寄出时间'].isnull())
                           & (~detail.loc[:,'预约上门时间'].isnull())
                           & (~detail.loc[:,'Fsender_phone'].isin(black_phone_list)) ]
    order_booking = order_booking1.copy()
    order_booking.loc[:,'预约时间间隔'] = (order_booking.loc[:,'预约上门时间'].apply(lambda x:pd.to_datetime(x,errors='coerce',format='%Y-%m-%d')).dt.date  - order_booking.loc[:,'下单时间'].dt.date ).dt.days  #转化为仅日期相减，忽略时间
    order_booking = order_booking[order_booking.loc[:,'预约时间间隔'] >= 2] #取日期相差2天的预约上门回收单
    
    sended_status_list = [ '已取消','已付款','议价中','待付款','待退货','待检测','已发货']
    # 根据条件过滤得 vivo渠道与投放H5，2日前下的回收订单,并且未寄出
    not_send_2daybefore = detail[(detail.loc[:,'下单时间'].apply(lambda x:x.strftime('%Y年%m月%d日')) >= the_2day_before)  #2日前下的回收订单
                           & ( detail.loc[:,'下单时间'] < (datetime.now() + timedelta(days=-1)))     # 超24小时
                           & (~detail.loc[:,'下单状态'].isin(sended_status_list))
                           & (detail.loc[:,'fchannel_id'].isin([10000121,10000943])) #筛选vivo渠道与投放H5两个渠道
                           & (detail.loc[:,'回收类型'].isin(recycle_type_list))
                           & (detail.loc[:,'业务类型'].isin(business_mode))
                           & (detail.loc[:,'寄出时间'].isnull()) # 未寄出
                           & (detail.loc[:,'收货时间'].isnull()) # 未收货，有异常缺失寄出的时间的情况
                           # & (~detail.loc[:,'预约上门时间'].isnull())
                           & (~detail.loc[:,'Fsender_phone'].isin(black_phone_list)) ].sort_values(by='main_channel')
    
    # 根据条件过滤得 除了vivo渠道与投放H5，3日前下的回收订单,并且未寄出
    not_send_3daybefore = detail[(detail.loc[:,'下单时间'].apply(lambda x:x.strftime('%Y年%m月%d日')) == the_3day_before)  #3日前下的回收订单
                           & ( detail.loc[:,'下单时间'] < (datetime.now() + timedelta(days=-1)))     # 超24小时
                           & (~detail.loc[:,'下单状态'].isin(sended_status_list))
                           & (~detail.loc[:,'fchannel_id'].isin([10000121,10000943])) #筛选除vivo渠道与投放H5两个渠道
                           & (detail.loc[:,'回收类型'].isin(recycle_type_list))
                           & (detail.loc[:,'业务类型'].isin(business_mode))
                           & (detail.loc[:,'寄出时间'].isnull()) # 未寄出
                           & (detail.loc[:,'收货时间'].isnull()) # 未收货，有异常缺失寄出的时间的情况                      
                           & (~detail.loc[:,'Fsender_phone'].isin(black_phone_list)) ].sort_values(by='main_channel')
    
    tmp1 = order_booking.copy()
    tmp2 = not_send_2daybefore.copy()
    tmp3 = not_send_3daybefore.copy()
    tmp1['来源'] = '预约时间过长'
    tmp2['来源'] = '2天前下单'
    tmp3['来源'] = '3天前下单'
    tmp = pd.concat([tmp1.loc[:,['来源','main_channel','Fsender_phone','用户注册手机号','寄件用户名','forder_id','预约上门时间','下单时间']],
                     tmp2.loc[:,['来源','main_channel','Fsender_phone','用户注册手机号','寄件用户名','forder_id','预约上门时间','下单时间']],
                     tmp3.loc[:,['来源','main_channel','Fsender_phone','用户注册手机号','寄件用户名','forder_id','预约上门时间','下单时间']]
                     ]).drop_duplicates(subset=['Fsender_phone'],keep='first', ignore_index= True)
    not_send_result = tmp.copy()
    not_send_result['外呼日期'] = today
    not_send_result_columns = ['来源','外呼日期','main_channel','forder_id','Fsender_phone','用户注册手机号','寄件用户名','预约上门时间','下单时间']
    
    
    # 构建智齿外呼机器人所需的传入外呼客户格式
    data_phone = not_send_result[['forder_id','Fsender_phone']].rename({'forder_id':'outId','Fsender_phone':'telephone'},axis=1)
    phone_json_tmp = data_phone.to_json(orient = 'records')
    phone_json = json.dumps( {'contactList':json.loads(phone_json_tmp)} )
    
    
    # 智齿提供的密钥
    apiKey = 'Gan8g1IAdQkWVR71KlbmlcWUvsC1m6hq'
    apiSecret = 's6sLAJzC4EpYp9h7FpTzJMlC7l0Nizff'
    companyId = '8eb4f95db03c46feb606e4da9f1b7f26'
    taskId = 'a6160b1cfca54e66a3c563514debfb04'
    
    # 调用自行编写的Sobotcall 模块
    getresult = Sobotcall.Sobot(apiKey,apiSecret) #实例化类
    getresult.getToken()
    getresult.addnum(companyId,taskId,phone_json)  #添加外呼数据
    
    
    # data_dict = {
    #     "contactList":[{
    #         "outId":"test001",
    #         "telephone":"17688938632",
    
    #     }]   
    # }
    # data = json.dumps(data_dict)
    # getresult.addnum('8eb4f95db03c46feb606e4da9f1b7f26','a6160b1cfca54e66a3c563514debfb04',data)
    
    # old_send = concat([send_push_bookding.loc[:,]
    #     ])
    
    print('开始导出所有结果')
    # file_path = r'E:\自有后端'
    file_path = father_record_file_path
    
    
    writer1 = pd.ExcelWriter(file_path+ '\\催寄记录\\' + '{}.xlsx'.format(today)) 
    not_send_result.to_excel(writer1,'{}'.format(today),index=False,columns=not_send_result_columns)
    writer1.save()
    writer1.handles = None
    print('已完成导出')

# def table_srt(dataframe):
#     tb = pt.PrettyTable()
#     tb.add_column('编号',dataframe.index)
#     for col in dataframe.columns.values:#df.columns.values的意思是获取列的名称
#         tb.add_column(col, dataframe[col])
#     print(tb)
#     return tb

# not_send_result_tb_str = str(table_srt(not_send_result[['来源','main_channel','Fsender_phone','用户注册手机号','寄件用户名','forder_id']]))

# url1 = 'https://oapi.dingtalk.com/robot/send?access_token=2ecee5fc7688fbbdca942cc1c2f656446edea37ac63a420e354b499a02f4b41e'
# messge1 = "一、{0}催寄-催寄\n".format(today) + not_send_result_tb_str
# reminders =[]
# print(send_msg(url1, reminders, messge1))
# etime = datetime.now()
# print("本次运行耗时%s秒" % (etime - stime).seconds)


if __name__ == '__main__':
    stime = datetime.now()
    # main()
    # 创建定时任务
    sched = BlockingScheduler()
    # 定时周一到周日每天9.20播报
    sched.add_job(main, 'cron', hour='11', minute='50', second='00',day_of_week='0-6',coalesce=True, misfire_grace_time=3600)
    sched.start()

