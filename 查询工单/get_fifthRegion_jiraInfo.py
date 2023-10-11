#encoding = utf-8
import requests, time, xlsxwriter, re, os
from pyquery import PyQuery as pq
from datetime import datetime

print("1、该脚本运行需要在电脑已经连接上公司的VPN的情况下运行！！\n2、因为接口请求的Cookie是会失效的，目前还没有找到有效办法自动获取Cookie，后面会想办法优化成自动填充，暂时请参考右边的链接中的方法来获取Cookie：https://alidocs.dingtalk.com/i/nodes/3QD5Ea7xAo4VEXQ373nKWG1YBwgnNKb0#！！\n3、并在下面的提示后面填上你拿到最新的Cookie，然后脚本即可正常运行！！\n")
time.sleep(5)

# StartTime = '2022-01-01'
# EndTime = '2022-12-31'
StartTime = input('请输入查询工单创建时间（开始，如 2022-01-01）：')
EndTime = input('请输入查询工单创建时间（结束，如 2022-12-31）：')
# Cookie = 'JSESSIONID=B49F6A28EDA4CC0B021B82B0990275CE; atlassian.xsrf.token=BJYK-W8TT-2ZEX-W82D_816e65efcfe02673f3ad5df99e318d95de94d3b3_lin'
Cookie = input('请输入Cookie：')


def get_gongdan_list(StartTime, EndTime, Cookie, gdlx, name):
    url = 'https://jira.megvii-inc.com/rest/issueNav/1/issueTable'
    header = {
        '__amdModuleName': 'jira/issue/utils/xsrf-token-header',
    'Accept': '*/*',
    'Accept-Encoding': 'gzip, deflate, br',
    'Accept-Language': 'zh-CN,zh;q=0.9',
    'Connection': 'keep-alive',
    'Content-Length': '145',
    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
    'Cookie': '%s'%Cookie,
    'Host': 'jira.megvii-inc.com',
    'Origin': 'https://jira.megvii-inc.com',
    'Referer': 'https://jira.megvii-inc.com/issues/?jql=project = %s AND created >= %s AND created <= %s AND reporter in (%s) ORDER BY created DESC'%(gdlx, StartTime, EndTime, name),
    'sec-ch-ua': '"Not?A_Brand";v="8", "Chromium";v="108", "Google Chrome";v="108"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
    'Sec-Fetch-Dest': 'empty',
    'Sec-Fetch-Mode': 'cors',
    'Sec-Fetch-Site': 'same-origin',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36',
    'X-Atlassian-Token': 'no-check',
    'X-Requested-With': 'XMLHttpRequest'
        }
    data = {
        'startIndex': 0,
        'jql': 'project = {} AND created >= {} AND created <= {} AND reporter in ({}) ORDER BY created DESC'.format(gdlx, StartTime, EndTime, name),
        'layoutKey': 'list-view'
        }
    tmp = requests.post(url=url, headers=header, data=data)
    format = tmp.json()
    tmp_gongdan_list = format["issueTable"]["issueKeys"]
########################################################
# 这里是用正则表达式找到名字对应的汉字
    if len(tmp_gongdan_list) > 0:
        tab = format["issueTable"]["table"]
        list = tab.split('\n')
        for k, v in enumerate(list):
            if v.rfind(name) != -1:
                a = v.rfind(name)
                sub_str = (v[a:a + 20])
                pre = re.compile(u'[\u4e00-\u9fa5]')   #正则表达式提取字符串的汉字部分
                res = re.findall(pre, sub_str)
                real_name_tmp = ''.join(res)
        real_name = real_name_tmp  
        # print(real_name)
######################################################
        return tmp_gongdan_list, real_name


def get_gongdan_info(gongdan_no, Cookie):
    info_list = []
    url = "https://jira.megvii-inc.com/browse/{}".format(gongdan_no)
    payload={}
    headers = {
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
    'Accept-Language': 'zh-CN,zh;q=0.9',
    'Cache-Control': 'max-age=0',
    'Connection': 'keep-alive',
    'Cookie': '%s'%Cookie,
    'Referer': 'https://jira.megvii-inc.com/issues/?jql=ORDER BY reated DESC',
    'Sec-Fetch-Dest': 'document',
    'Sec-Fetch-Mode': 'navigate',
    'Sec-Fetch-Site': 'same-origin',
    'Sec-Fetch-User': '?1',
    'Upgrade-Insecure-Requests': '1',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36',
    'sec-ch-ua': '"Not?A_Brand";v="8", "Chromium";v="108", "Google Chrome";v="108"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"'
    }

    response = requests.request("GET", url, headers=headers, params=payload)
    doc = pq(response.text)
    # doc = pq(url=url, headers=headers, encoding="utf-8")
    gongdan_creater = doc('span#reporter-val').text()
    gongdan_name = doc('title').text().replace(" - Your Company Jira", "")   #将字符串中多出来的 - Your Company Jira给替换成空
    project_name = doc('li#rowForcustomfield_11632 div.wrap div#customfield_11632-val').text()
    if len(project_name) == 0:
        project_name = "无项目名称"     #判断字符串是否为空的方法，如果字符串由纯空格组成可以用str.isspace()判断
    project_no = doc('li.item div.wrap div#customfield_11503-val').text()
    if len(project_no) == 0:
        project_no = "无项目编号"
    jbr = doc('span#assignee-val').text()   #经办人
    createDate = doc('span[data-name="创建日期"] time.livestamp').attr("datetime").replace("T", " ").replace("+0800", "")   #创建日期，doc().attr("元素") 返回所选元素的值
    reslution_result = doc('span#resolution-val').text()   #解决结果
    reslutionDate = doc('span#resolutiondate-val time.livestamp').attr("datetime")  #已解决（时间→空，就表示未解决）
    if reslutionDate is None:
        reslutionDate = "None"
        hour_sub = "None"   #解决耗时
    else:
        reslutionDate = reslutionDate.replace("T", " ").replace("+0800", "")
        cjsj = datetime.strptime(createDate, "%Y-%m-%d %H:%M:%S")
        jjsj = datetime.strptime(reslutionDate, "%Y-%m-%d %H:%M:%S")
        total_seconds = (jjsj - cjsj).total_seconds()
        hour_sub = total_seconds / 3600
        hour_sub = '{:.2f}'.format(hour_sub)  #解决耗时

    info_list.append(gongdan_creater)
    info_list.append(gongdan_name)
    info_list.append(project_name)
    info_list.append(project_no)
    info_list.append(jbr)
    info_list.append(createDate)
    info_list.append(reslution_result)
    info_list.append(reslutionDate)
    info_list.append(hour_sub)

    link_gongdan_tmp = doc('div.links-container p span a[class="issue-link link-title resolution"]').text()

    find_Commit = link_gongdan_tmp.find("Commit")
    if len(link_gongdan_tmp) != 0:
        if find_Commit != -1:
            len_gd_tmp = len(link_gongdan_tmp)
            jiequ_str = link_gongdan_tmp[find_Commit:len_gd_tmp]
            link_gongdan = link_gongdan_tmp.replace(jiequ_str, "")
            info_list.append(link_gongdan)
        else:
            info_list.append(link_gongdan_tmp.split(" "))
    elif len(link_gongdan_tmp) == 0:
        link_gongdan = "无关联工单"
        info_list.append(link_gongdan)
    return info_list
    '''CSS选择器参考样例，标签之间用空格（什么都不加），id的话用#号，class就用.号, a[class="issue-link link-title resolution"]选择指定key内容
    <li id="rowForcustomfield_11632" class="item">
        <div class="wrap">
            <strong title="项目名称" class="name">项目名称:</strong>
            <div id="customfield_11632-val" class="value type-textfield" data-fieldtype="textfield" data-fieldtypecompletekey="com.atlassian.jira.plugin.system.customfieldtypes:textfield">
                                          北京市公安局视综平台新建2022项目
  
                            </div>
        </div>
    </li>'''


def get_linkgongdan_info(link_gd_no, Cookie):
    url = "https://jira.megvii-inc.com/browse/{}".format(link_gd_no)
    payload={}
    headers = {
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
    'Accept-Language': 'zh-CN,zh;q=0.9',
    'Cache-Control': 'max-age=0',
    'Connection': 'keep-alive',
    'Cookie': '%s'%Cookie,
    'Referer': 'https://jira.megvii-inc.com/issues/?jql=ORDER BY reated DESC',
    'Sec-Fetch-Dest': 'document',
    'Sec-Fetch-Mode': 'navigate',
    'Sec-Fetch-Site': 'same-origin',
    'Sec-Fetch-User': '?1',
    'Upgrade-Insecure-Requests': '1',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36',
    'sec-ch-ua': '"Not?A_Brand";v="8", "Chromium";v="108", "Google Chrome";v="108"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"'
    }

    link_res = requests.request("GET", url, headers=headers, params=payload)
    link_doc = pq(link_res.text)
    link_gongdan_creater = link_doc('span#reporter-val').text()
    link_jbr = link_doc('span#assignee-val').text()   #经办人
    link_createDate = link_doc('span[data-name="创建日期"] time.livestamp').attr("datetime").replace("T", " ").replace("+0800", "")   #创建日期，doc().attr("元素") 返回所选元素的值
    # link_reslution_result = link_doc('span#resolution-val').text()   #解决结果
    link_reslutionDate = link_doc('span#resolutiondate-val time.livestamp').attr("datetime")  #已解决（时间→空，就表示未解决）
    if link_reslutionDate is None:
        link_reslutionDate = "None"
        link_hour_sub = "未解决"   #解决耗时
    else:
        link_reslutionDate = link_reslutionDate.replace("T", " ").replace("+0800", "")
        cjsj = datetime.strptime(link_createDate, "%Y-%m-%d %H:%M:%S")
        jjsj = datetime.strptime(link_reslutionDate, "%Y-%m-%d %H:%M:%S")
        total_seconds = (jjsj - cjsj).total_seconds()
        link_hour_sub = total_seconds / 3600
        link_hour_sub = '{:.2f}'.format(link_hour_sub)  #解决耗时
    return link_gongdan_creater, link_jbr, link_hour_sub



#结果写入excel文件
# wb = xlsxwriter.Workbook(EndTime[0:4] + "年度工单数量.xls")  #创建工作簿
dir = os.getcwd()
# file_name = f'五组{StartTime}~{EndTime}工单分布情况.xls'
file_name = "五组{}~{}工单分布情况.xls".format(StartTime, EndTime)   #两种字符串格式化处理的方法
full_path = os.path.join(dir, file_name)
# print('\n五组{}~{}工单分布情况见excel文件： {}'.format(StartTime, EndTime, full_path))
wb = xlsxwriter.Workbook(full_path)  #创建工作簿
sheet = wb.add_worksheet("五组{}~{}工单分布情况".format(StartTime, EndTime))  #创建工作表
sheet_1 = wb.add_worksheet("五组{}~{}工单分布详情".format(StartTime, EndTime))  #创建工作表
cell_format1 = wb.add_format()   #通过创建好的格式对象，直接调用对应的方法
cell_format1.set_bold()  #设置加粗，默认加粗
cell_format1.set_font_color("blue")  #文本颜色
cell_format1.set_font_size(18)      #设置字号
cell_format1.set_align("center")   #左右居中
cell_format1.set_align("vcenter")  #上下居中
# cell_format1.set_bg_color("gray")  #设置背景色
# cell_format1.set_border(style=1)   #设置边框格式
sheet.merge_range(0, 0, 1, 7, "五组{}~{}工单分布情况".format(StartTime, EndTime), cell_format1)

#格式对象2
cell_format2 = wb.add_format()
cell_format2.set_border(style=1)
cell_format2.set_bold()  
first_row = ['姓名', 'GDCSP', 'SSTSP', 'CORESP', 'IOTSP', 'PLATFORMSP', 'CBGRMS', '个人合计']
cell_format2.set_bg_color("gray")  #设置背景色
sheet.write_row(2, 0, first_row, cell_format2)

#格式对象3
cell_format3 = wb.add_format()
cell_format3.set_border(style=1)

#格式对象4
cell_format4 = wb.add_format()
cell_format4.set_border(style=1)
cell_format4.set_bold()  
first_row_4 = ['报告人', '工单号及工单名称', '项目名称', '项目编号', '经办人', '创建日期', '解决结果', '解决日期', '解决耗时(小时)', '关联工单号', '关联工单报告人', '关联工单经办人', '关联工单解决耗时(小时)']
cell_format4.set_bg_color("gray")  #设置背景色
sheet_1.write_row(0, 0, first_row_4, cell_format4)



if __name__ == "__main__":
    try:
        count_title = ['姓名', 'GDCSP', 'SSTSP', 'CORESP', 'IOTSP', 'PLATFORMSP', 'CBGRMS']
        print(count_title)
        gongdan_sum = 0
        gongdan_sum_list=[]
        # name_list = ['wangbo02', 'wanxiaoxiang', 'fanjialong', 'zhangwei03', 'linshugui', 'wujiaqi', 'dushuangqing']
        name_dict = {'zhangmin03':'张敏', 'wanxiaoxiang':'宛小祥', 'hemengmeng':'何蒙蒙', 'yangxin04':'杨新', 'weirongci':'魏榕祠', 'huguanghui':'胡光辉'}
        gdlx_list = ['GDCSP', 'SSTSP', 'CORESP', 'IOTSP', 'PLATFORMSP', 'CBGRMS']
        rows_of_gd_count = []
        finally_gongdan_info = []

        for name in name_dict:
            everyone_count = []
            for gdlx in gdlx_list:
                gongdan_list = get_gongdan_list(StartTime, EndTime, Cookie, gdlx, name)
                if gongdan_list is not None:
                    # for gd_bh in gongdan_list[0]:
                    #     print(gd_bh)
                    per_count = len(gongdan_list[0]) 
                    # print(gongdan_list[1] + " 2022年度的" + gdlx + "工单总数：" + str(len(gongdan_list)))
                    gongdan_sum+=per_count
                    everyone_count.append(gdlx)
                    everyone_count.append(str(per_count))

                    for gongdan_no in gongdan_list[0]:
                        gongdan_info = get_gongdan_info(gongdan_no, Cookie)
                        # print(gongdan_info)
                        link_gd_list = gongdan_info[-1]
                        # print(link_gd_list)
                        sub_gongdan_info = gongdan_info[0:-1]
                        # print(sub_gongdan_info)
                        if link_gd_list != "无关联工单":
                            for link_gd_no in link_gd_list:
                                link_gd_result = get_linkgongdan_info(link_gd_no, Cookie)
                                link_gd_cjr = link_gd_result[0]
                                link_gd_jbr = link_gd_result[1]
                                link_gd_useTime = link_gd_result[2]
                                sub_gongdan_info.append(link_gd_no)
                                sub_gongdan_info.append(link_gd_cjr)
                                sub_gongdan_info.append(link_gd_jbr)
                                sub_gongdan_info.append(link_gd_useTime)
                            finally_gongdan_info.append(sub_gongdan_info)
                            # print(sub_gongdan_info)
                        else:
                            sub_gongdan_info.append("无关联工单")
                            finally_gongdan_info.append(sub_gongdan_info)
                    # print(finally_gongdan_info)
                else:
                    everyone_count.append(gdlx)
                    everyone_count.append(str(0))               
                    
            everyone_count.insert(0, name_dict[name])
            finaly_count_list_tmp = everyone_count[::2]    #取列表索引为偶数的值，奇数的为everyone_count[1::2]
            print(finaly_count_list_tmp)    #输出各成员的工单数量
            rows_of_gd_count.append(finaly_count_list_tmp)

        GDCSP_SUM = 0
        SSTSP_SUM = 0
        CORESP_SUM = 0 
        IOTSP_SUM = 0
        PLATFORMSP_SUM = 0
        CBGRMS_SUM = 0
        ALL_SUM = 0
        new_row = ['分类合计']
        for x, y in enumerate(rows_of_gd_count):
            y[1:] = map(int, y[1:])
            rows_of_gd_count[x].append(sum(y[1:]))
            y[1:] = map(str, y[1:])
            # print(x, y)
            gdc_count = rows_of_gd_count[x][1]
            # print(gdc_count)
            GDCSP_SUM+=int(gdc_count)
            sstsp_count = rows_of_gd_count[x][2]
            SSTSP_SUM+=int(sstsp_count)
            coresp_count = rows_of_gd_count[x][3]
            CORESP_SUM+=int(coresp_count)
            iotsp_count = rows_of_gd_count[x][4]
            IOTSP_SUM+=int(iotsp_count)
            platformsp_count = rows_of_gd_count[x][5]
            PLATFORMSP_SUM+=int(platformsp_count)
            cbgrms_count = rows_of_gd_count[x][6]
            CBGRMS_SUM+=int(cbgrms_count)
            all_count = rows_of_gd_count[x][7]
            ALL_SUM+=int(all_count)

        new_row.append(GDCSP_SUM)
        new_row.append(SSTSP_SUM)
        new_row.append(CORESP_SUM)
        new_row.append(IOTSP_SUM)
        new_row.append(PLATFORMSP_SUM)
        new_row.append(CBGRMS_SUM)
        new_row.append(ALL_SUM)
        rows_of_gd_count.append(new_row)

        for i, v in enumerate(rows_of_gd_count):
            v[1:] = map(str, v[1:])
        # print(rows_of_gd_count)

        for index, item in enumerate(rows_of_gd_count):    #一行一行循环写入   
            sheet.write_row(index + 3, 0, item, cell_format3)
        for k, v in enumerate(finally_gongdan_info):
            # print(k, v)
            sheet_1.write_row(k + 1, 0, v, cell_format3)
        print("五组人员提出工单总数：" + str(gongdan_sum))
        print('\n五组{}~{}工单分布情况见excel文件： {}'.format(StartTime, EndTime, full_path))
        wb.close()
        time.sleep(120)
    except:
        print('异常、主动退出或无法获取数据，请确认VPN是否连接，Cookie是否过期！！')
        time.sleep(60)
