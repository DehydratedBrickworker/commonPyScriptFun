# Rpa开发常用的python功能代码

# 创建文件或文件夹的功能
def mkDirOrFile(path:str, isrecreate:bool=True, mkwhat:str=None, sheet_name=None):
    '''
    :param path:
    :param isrecreate: recreate (If the file already exists) ?
    :param mkwhat: make file or folder?
    :param sheet_name: sheet name for create excel book
    :return:
    '''
    import os
    import shutil
    import openpyxl
    class disunityPathSlash(Exception):
        def __str__(self):
            return 'Unify the use of slashes or backslashes in the path'

    class srcFileNotExist(Exception):
        def __init__(self, srcfile):
            self.srcfile = srcfile

        def __str__(self):
            return f'source file <{self.srcfile}> not exist,Unable to copy'

    class parmTypeError(Exception):
        def __init__(self, varname):
            self.varname = varname

        def __str__(self):
            return f'parameter \'{self.varname}\' type or content not satisfiable，should refer to <dict_parmTypeJudge>'

    dict_parmTypeJudge = {
        'path' : type(path) in [str, list],
        'isrecreate' : type(isrecreate) is bool,
        'mkwhat' : mkwhat in ['Dir', 'File', 'Copy', 'Excel', None],
        'sheet_name' : (type(sheet_name) is dict) | (sheet_name is None),
    }

    for parm in dict_parmTypeJudge.keys():
        if dict_parmTypeJudge[parm] == False:
            raise parmTypeError(parm)

    def getFileName(path):
        '''
        get file or folder name
        :param path: file or folder path
        :return: file or folder name
        '''

        if ('\\' in path) & ('/' not in path):
            fileOrDirName = path.split('\\')[-1]
        elif ('\\' not in path) & ('/' in path):
            fileOrDirName = path.split('/')[-1]
        else:
            raise disunityPathSlash

        return fileOrDirName

    def pathExistAndHandleWay(path):
        '''
        print('{} file or folder already exist -> no need to create or copy'.format(getFileName(path)))
        print('{} file or folder already exist -> preparing to delete and recreate or copy'.format(getFileName(path)))
        print('{} file or folder already deleted and recreated or copied'.format(getFileName(path)))
        print('{} file or folder does not exist -> waiting for creation or copy'.format(getFileName(path)))
        print('{} file or folder already created or copied'.format(getFileName(path)))
        :return:
        '''

        if os.path.exists(path):
            if isrecreate is False:
                print('{} file or folder already exist -> no need to create or copy'.format(getFileName(path)))
                howToDo = None
            elif isrecreate is True:
                print('{} file or folder already exist -> delete and recreate or copy'.format(getFileName(path)))
                howToDo = 'Recreate'
        else:
            print('{} file or folder does not exist -> creation or copy'.format(getFileName(path)))
            howToDo = 'Create'

        return howToDo

    def mkDir():
        howToDo = pathExistAndHandleWay(path)
        if howToDo is None:
            pass
        elif howToDo == 'Recreate':
            shutil.rmtree(path)
            os.mkdir(path)
        elif howToDo == 'Create':
            os.mkdir(path)

    def mkFile():
        howToDo = pathExistAndHandleWay(path)
        if howToDo is None:
            pass
        elif howToDo == 'Recreate':
            os.remove(path)
            file = open(path, 'w')
            file.close()
        elif howToDo == 'Create':
            file = open(path, 'w')
            file.close()

    def copyFile():
        howToDo = pathExistAndHandleWay(path[1])
        if os.path.exists(path[0]):
            if howToDo is None:
                pass
            elif howToDo == 'Recreate':
                os.remove(path[1])
                shutil.copyfile(path[0], path[1])
            elif howToDo == 'Created':
                shutil.copyfile(path[0], path[1])
        else:
            raise srcFileNotExist(path[0])

    def mkExcel():

        def mkExcel():

            workbook = openpyxl.Workbook(path)
            if sheet_name is not None:
                for sht_loc in sheet_name.keys():
                    worksheet = workbook.create_sheet(sht_loc)
                    worksheet.title = sheet_name[sht_loc]
            workbook.save(path)

        howToDo = pathExistAndHandleWay(path)
        if howToDo is None:
            pass
        elif howToDo == 'Recreate':
            os.remove(path)
            mkExcel()
        elif howToDo == 'Create':
            mkExcel()


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

    is_XlsxOrXlsm = '.' + path.split('\\')[-1].split('.')[-1] in ['xlsx', '.xlsm']
    if type(path) is str:
        if '.' + path.split('\\')[-1].split('.')[-1] in list_FileType:
            if is_XlsxOrXlsm:
                isWhat = 'Excel'
            elif ('.' + path.split('\\')[-1].split('.')[-1] in list_FileType) & (not is_XlsxOrXlsm):
                isWhat = 'File'
        else:
            isWhat = 'Dir'
    elif type(path) is list:
        isWhat = 'Copy'

    if (mkwhat == 'Dir') | (isWhat == 'Dir'):
        mkDir()
    elif (mkwhat == 'File') | (isWhat == 'File'):
        mkFile()
    elif (mkwhat == 'Copy') | (isWhat == 'Copy'):
        copyFile()
    elif (mkwhat == 'Excel') | (isWhat == 'Excel'):
        mkExcel()

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

    class parmTypeError(Exception):
        def __init__(self, varname):
            self.varname = varname

        def __str__(self):
            return f'传入参数\'{self.varname}\'的类型或内容不符合要求，查看内部函数dict_parmTypeJudge'

    dict_parmTypeJudge = {
        'path' : type(path) is str,
        'istranspose' : type(istranspose) is bool,
        'isvisible' : type(isvisible) is bool,
        'isaddbook' : type(isaddbook) is bool,
        'sheetname' : type(sheetname) in [str, int],
        'values' : type(values) is dict,
        'columns' : (type(columns) is list) & (columns is None),
    }

    for parm in dict_parmTypeJudge.keys():
        if dict_parmTypeJudge[parm] == False:
            raise parmTypeError(parm)

    app = xw.App(visible=isvisible, add_book=isaddbook)
    wb = app.books.open(path)
    sht = wb.sheets[sheetname]
    print('开始写入Excel')
    if columns is not None:
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
