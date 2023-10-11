from pymongo import MongoClient
from colorama import init,Fore
import pandas as pd
import pymysql
import time
import sys


init(autoreset=True)

class getAlbumUpdateStatus(object):

    def __init__(self, ip, devops_infra_passwd, needIndex = 'createTime', dkId = 'albumId'):
        self.ip = ip
        self.devops_infra_passwd = devops_infra_passwd
        self.url = f'mongodb://adminx:{self.devops_infra_passwd}@{self.ip}:27027/admin'
        self.client = MongoClient(self.url)
        self.needIndex = needIndex
        self.dkId = dkId
        self.albumInfo = []
        self.picInfo = []

    def countdown(self):
        total_time = 30
        while total_time >= 0:         
            minutes, seconds = divmod(total_time, 60)      # 将秒钟数转换为分钟和秒钟（用divmod()函数实现）
            time_format = '{:02d}:{:02d}'.format(minutes, seconds)   # 定义时间格式
            print(time_format, end=Fore.RED + '\r程序退出倒计时：')     # 输出当前时间，为实现动态效果，需要加'\r'字符
            total_time -= 1      # 总时间减少1秒钟
            time.sleep(1)       # 暂停1秒钟, 防止CPU占用率过高

    #查询底库创建人的用户名，并更新self.albumInfo里面的创建人userid为user_name
    def getUserName(self):
        db = pymysql.connect(host=self.ip, port=3306, user='root', password=self.devops_infra_passwd, database='auth2', charset='utf8')
        cursor = db.cursor()
        for userId in self.albumInfo:
            sql = f"select user_name from auth_user where id='{userId[6]}';"
            cursor.execute(sql)
            userName = cursor.fetchall()
            userId[6] = userName[0][0]

    def getAlbumNameInfo(self):
        db = self.client['picture']
        fc_album = db['fc_album']
        albumQuery = fc_album.find({}, {'_id':1, 'name':1,'clct.photo':1, 'create_time': 1, 'type': 1, 'creator': 1, 'auth.type': 1, 'album_list.device': 1, 'property.hasRealname': 1})
        for j in albumQuery:
            eachAlbum = []
            for k, v in j.items():
                if isinstance(v, dict):
                    for x, y in v.items():
                        # print(y)
                        eachAlbum.append(y)
                elif isinstance(v, list):
                    if not v[0]:
                        # print('Null')
                        eachAlbum.append('Null')
                    else:
                        for g in v:
                            for p, q in g.items():
                                # print(q)
                                eachAlbum.append(q)
                else:
                    # print(v)
                    eachAlbum.append(v)
            # print(eachAlbum)
            self.albumInfo.append(eachAlbum)

        #对已经生成的列表self.albumInfo里面的值进行转换、更新，如：时间戳、字典值
        for each in self.albumInfo:
            #将底库的创建时间戳转换成时间
            albumCreateStructTime = time.localtime(each[7] / 1000)
            timeConvert = time.strftime('%Y-%m-%d %H:%M:%S', albumCreateStructTime)
            each[7] = timeConvert
            #将是否实名化的字典值，翻译成映射中文
            if len(each) < 9:
                each.append('未开启')
            elif each[8] != '1':
                each[8] = '未开启'
            else:
                each[8] = '开启'
            #将底库类型的字典值转换成映射的中文值
            if each[3] == 1:
                each[3] = '静态库'
            else:
                each[3] = '布控库'
            #将底库权限的字典值转换成映射的中文值
            if each[4] == 0:
                each[4] = '私有库'
            elif each[4] == 1:
                each[4] = '自定义'
            else:
                each[4] = '公共库'

    def createIndex(self):
        db = self.client['picture']
        AlbumNameInfo = {yy[2]:yy[5] for yy in self.albumInfo}    #将self.albumInfo里面的底库名称和底库表明组成字典，以作后用
        for albumName, dbName in AlbumNameInfo.items():
            # print(name, dbName)
            if dbName != '无':
                dbNameM = db[dbName]
                all_index_string = ''
                all_index = dbNameM.list_indexes()
                for eachIndexRow in all_index:
                    all_index_string += str(eachIndexRow)
                if all_index_string.find(self.needIndex) == -1:
                    dbNameM.create_index(self.needIndex)
                    print(Fore.MAGENTA + f"已为底库《{albumName}》, 集合《{dbName}》创建索引：{self.needIndex};")
                else:
                    print(Fore.GREEN + f"底库《{albumName}》, 集合《{dbName}》已存在索引：{self.needIndex};")

    def queryFirstAndLast(self):
        db = self.client['picture']
        AlbumNameInfo = {yy[2]:yy[5] for yy in self.albumInfo}    #将self.albumInfo里面的底库名称和底库表明组成字典，以作后用
        for albumName, dbName in AlbumNameInfo.items():
            if dbName != '无':
                everyTwo = []
                dbNameM = db[dbName]   
                albumCount = dbNameM.count_documents({})
                if albumCount > 0:
                    queryLatest = dbNameM.find({},{'_id': 0, self.dkId: 1, self.needIndex: 1}).sort(self.needIndex, 1).limit(1)
                    for eachLatest in queryLatest:
                        structTimeNew = time.localtime(eachLatest[self.needIndex] / 1000)
                        usualTimeNew = time.strftime('%Y-%m-%d %H:%M:%S', structTimeNew)
                        print(Fore.CYAN + f"底库《{albumName}》最旧的照片创建时间：{usualTimeNew}")
                        everyTwo.append(eachLatest[self.dkId])
                        everyTwo.append(albumName)
                        everyTwo.append(albumCount)
                        everyTwo.append(usualTimeNew)
                    queryOld = dbNameM.find({},{self.needIndex: 1,'_id': 0}).sort(self.needIndex, -1).limit(1)
                    for eachOld in queryOld:
                        structTimeOld = time.localtime(eachOld[self.needIndex] / 1000)
                        usualTimeOld = time.strftime('%Y-%m-%d %H:%M:%S', structTimeOld)
                        print(Fore.CYAN + f"底库《{albumName}》最新的照片创建时间：{usualTimeOld}")
                        everyTwo.append(usualTimeOld)
                # print(everyTwo)
                if len(everyTwo) > 0:
                    self.picInfo.append(everyTwo)


def main():
    ip = input("请输入业务IP：")
    # ip = '32.109.131.130'
    devops_infra_passwd = input("请输入devops环境变量：")
    # devops_infra_passwd = 'PwIZ9i6dgef'
    needIndex = input('请输入底库表中照片创建时间的字段名称，不输入则默认是(createTime)：')
    dkId = input('请输入底库表中底库ID的字段名称，不输入则默认是(albumId)：')
    if needIndex == "" and dkId == "":
        call = getAlbumUpdateStatus(ip, devops_infra_passwd)
    else:
        call = getAlbumUpdateStatus(ip, devops_infra_passwd, needIndex, dkId)

    call.getAlbumNameInfo()
    call.getUserName()
    #对fc_album里面查询到的底库信息进行数据二维化
    df_fa = pd.DataFrame(call.albumInfo, columns=['底库ID', '底库device类型', '底库名称', '底库类型', '底库权限', '底库表名', '创建人', '底库创建时间', '是否参与实名化'])
    col_list = ['底库ID', '底库device类型', '底库名称', '底库表名', '底库类型', '底库权限', '创建人', '底库创建时间', '是否参与实名化']  #调整列底库表名到第四列
    df_fa = df_fa[col_list]
    # df_fa.set_index('底库ID', inplace=True)
    # print(df_fa.head())
    # df_fa.to_excel(r'C:\Users\74049\Desktop\fc_album_info.xlsx')

    #为创建时间createTime创建索引，如果已存在则不创建，其实mongo本身也会有判断，我们这里不做判断也是可以的
    call.createIndex()

    #遍历每个库查看最新数据的创建时间和最旧数据的创建时间来确认底库是否有新增照片
    call.queryFirstAndLast()
    #将查询的照片信息进行数据二维化
    df_pic = pd.DataFrame(call.picInfo, columns=['底库ID', '底库名称', '底库人像数', '底库最旧人像创建时间', '底库最新人像创建时间'])
    # df_pic.set_index('底库ID', inplace=True)
    # print(df_pic.head())
    # df_pic.to_excel(r'C:\Users\74049\Desktop\fc_album_info-1.xlsx')

    #merge聚合两组二位数据，相当于excel的vlookup的扩展
    mergeAll = pd.merge(df_fa, df_pic, how='left', on=['底库名称'])
    # print(mergeAll.head())
    # mergeAll.set_index('底库ID_X', inplace=True)
    mergeAll.index = mergeAll.index + 1
    mergeAll.drop(columns=['底库ID_y'], inplace=True)
    mergeAll = mergeAll.rename(columns={'底库ID_x': '底库ID'})
    mergeAll.to_excel(f"./底库更新情况_{time.strftime('%Y%m%d%H%M%S',time.localtime())}.xlsx", sheet_name="查看底库更新情况")

    #到这里则已经执行完成，下面会判断平台的类型来确认是否立即退出
    try:
        if sys.platform == 'win32':
            call.countdown()
    except KeyboardInterrupt:
        print('\nctrl+c 主动终止！')
    finally:    
        sys.exit(0)


if __name__ == "__main__":
    startTime = time.time()
    main()
    endTime = time.time()
    print(Fore.WHITE + f"程序执行耗时：{endTime - startTime}s")
