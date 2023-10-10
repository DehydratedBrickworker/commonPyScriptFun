# Rpa开发常用的python功能代码

# 创建文件或文件夹的功能
def mkDirOrFile(path:str, isrecreate:bool=True, mkwhat:str=None):
    '''
    :param path: 路径
    :param isrecreate: 如果文件已存在，是否删除并重新创建
    :param mkwhat: 创建文件还是文件夹
    :return:
    '''
    import os
    import shutil

    class parmTypeError(Exception):
        def __init__(self, varname):
            self.varname = varname

        def __str__(self):
            return f'传入参数\'{self.varname}\'的类型不正确'

    dict_parmTypeJudge = {
        'path' : type(path) is str,
        'isrecreate' : type(isrecreate) is bool,
        'mkwhat' : mkwhat in ['Dir', 'File'],
    }

    for parm in dict_parmTypeJudge.keys():
        if dict_parmTypeJudge[parm] == False:
            raise parmTypeError(parm)

    def mkDir(path):
        if os.path.exists(path):
            if isrecreate == False:
                print('{}文件夹已存在 -> 无需创建'.format(path.split('\\')[-1]))
            elif isrecreate == True:
                print('{}文件夹存在 -> 准备删除并重新创建'.format(path.split('\\')[-1]))
                shutil.rmtree(path)
                os.mkdir(path)
                print('{}文件夹已删除并重新创建'.format(path.split('\\')[-1]))
        else:
            print('{}文件夹不存在->等待创建'.format(path.split('\\')[-1]))
            os.mkdir(path)
            print('{}文件夹已创建'.format(path.split('\\')[-1]))

    def mkFile(path):
        if os.path.exists(path):
            if isrecreate == False:
                print('{}文件已存在 -> 无需创建'.format(path.split('\\')[-1]))
            elif isrecreate == True:
                print('{}文件存在 -> 准备删除并重新创建'.format(path.split('\\')[-1]))
                os.remove(path)
                file = open(path, 'w')
                file.close()
                print('{}文件已删除并重新创建'.format(path.split('\\')[-1]))
        else:
            print('{}文件不存在->等待创建'.format(path.split('\\')[-1]))
            file = open(path, 'w')
            file.close()
            print('{}文件已创建'.format(path.split('\\')[-1]))

    dict_fileType = {
        'Excel文件类型' : ['.xls', '.xlsx', '.csv', '.xlsm', '.et'],
        '文本文件类型' : ['.txt', '.doc', '.docx'],
        '音频文件类型' : ['.mp3', '.wav', '.aac'],
        '视频文件类型' : ['.mp4', '.wmv', '.avi', '.mov'],
        '图像文件类型' : ['.jpg', '.jpeg', '.png'],
        '压缩文件类型' : ['.zip', '.rar', '.7z'],
        '数据库文件类型' : ['.mdb', '.sqlite', '.db'],
        'PDF文件类型' : ['.pdf'],
        'PPT文件类型' : ['.ppt', '.pptx'],
    }

    list_FileType = []

    for filetype in dict_fileType.keys():
        list_FileType += dict_fileType[filetype]
    if '.' + path.split('\\')[-1].split('.')[-1] not in list_FileType:
        isFile = False
    else:
        isFile = True

    if (mkwhat == 'Dir') | (isFile == False):
        mkDir(path)
    elif (mkwhat == 'File') | (isFile == True):
        mkFile(path)

    return path


# 写入Excel的功能
def writeToExcel(path:str, values:dict, sheetname=0, columns:list=None, istranspose=False, isvisible=False, isaddbook=False):
    '''

    :param path: Excel文件路径
    :param values: 写入位置:写入的值 组成的字典
    :param sheetname: 页名
    :param columns: 列名组成的列表，默认值为None->跳过写入列名的步骤
    :param istranspose: 是否转置
    :return:
    '''
    import xlwings as xw

    app = xw.App(visible=isvisible, add_book=isaddbook)
    wb = app.books.open(path)
    sht = wb.sheets[sheetname]
    print('开始写入Excel')
    if (columns is not None) & (len(columns) > 0):
        sht.range('A1').value=columns
        print('成功写入列名')
    if len(values.keys()) > 0:
        for location in values.keys():
            sht.range(location).options(transpose=istranspose).value=values[location]
            print(f'成功在{location}写入数据')
    else:
        print('无写入位置及写入值')
    wb.save()
    wb.close()
    app.quit()


# 日期相关计算
import datetime

class dateTimeAbout:

    def __init__(self, dateTime):
        '''
        处理传入的日期时间
        :param dateTime: 日期时间，可以为字符串或时间格式
        '''
        if type(dateTime) is str:
            self.dateTime = dateTime.replace('.', '').replace('-', '').replace('/', '').replace(' ','')
            if len(self.dateTime) == 8:
                self.dateTime = datetime.datetime.strptime(self.dateTime, '%Y%m%d')
            elif len(self.dateTime) == 12:
                self.dateTime = datetime.datetime.strptime(self.dateTime, '%Y%m%d%H%M')
            elif len(self.dateTime) == 14:
                self.dateTime = datetime.datetime.strptime(self.dateTime, '%Y%m%d%H%M%S')
        else:
            self.dateTime = dateTime

    # 当前时间
    def nowDateTime(self):
        return self.dateTime

    # 前一天
    def lastDay(self):
        lastDay = self.dateTime + datetime.timedelta(days=-1)
        return lastDay

    # 当前日期上月的第一天
    def lastMonthFirstDay(self):
        if self.dateTime.month == 1:
            lastMonthFirstDay = datetime.datetime(self.dateTime.year - 1, 12 ,1)
        else:
            lastMonthFirstDay = datetime.datetime(self.dateTime.year, self.dateTime.month - 1, 1)
        return lastMonthFirstDay

    # 当前日期对应的星期
    def nowWeek(self):
        # 格式化输出当前日期时间
        t = self.dateTime.strftime('%Y%m%d%H%M%S')
        Weekday_Beijing = (datetime.datetime.strptime(t,'%Y%m%d%H%M%S') + datetime.timedelta(days=1)).weekday()
        Weekday_Greenwich = datetime.datetime.utcnow().weekday()
        return Weekday_Beijing
