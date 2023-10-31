# Rpa开发常用的python功能代码

# Excel表格及其内容是否存在检测的功能
def fileAbsentCheck(path):
    '''
    table and content is exists?
    :param path: table path
    :return: bool : table and content is exists?
    '''
    import os
    import pandas as pd

    contentIsNone = True
    tableIsExist = os.path.exists(path)
    if tableIsExist:
        if os.path.getsize(path) > 0:
            df_Table = pd.read_excel(path)
            if df_Table.shape[0] > 0:
                tableIsExist = True
            else:
                tableIsExist = False
        else:
            tableIsExist = False

    if (tableIsExist is True) & (contentIsNone is True):
        print(f'{path} table and content is exists')
        return True
    else:
        print(f'{path} table or content not exists')
        return False


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

    :param path: path of Excel file
    :param values: dict of table locate:value
    :param sheetname: pagename
    :param columns: list of table columns
    :param istranspose:
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

def dfDiffContentHandle(df:list, connect_condition, find_what:dict, sec_table_sync=False):
    '''
    两个通过具有唯一性的查询条件，比较各字典是否不同
    :param df: list [主数据框, 副数据框]
    :param connect_condition: list or str [查询条件字段] or [主数据框查询条件字段, 副数据框查询条件字段]
    :param find_what: dict 需要查询的内容是否不同的字段的字典 {主数据框字段 : 对应主数据框字段的副数据框字段}
    :param sec_table_sync: bool 是否需要按照副数据框字段的内容更改主数据框字段的内容
    :return:
    '''
    import pandas

    class parmTypeError(Exception):
        def __init__(self, varname):
            self.varname = varname

    dict_parmTypeJudge = {
        'df' : (type(df) is list) & (False not in [type(x) is pandas.core.frame.DataFrame for x in df]),
        'connect_condition' : (type(connect_condition) is str) | \
            ((type(connect_condition) is list) & (False not in [type(x) is str for x in connect_condition if type(connect_condition) is list]) & (len(connect_condition) == 2 if type(connect_condition) is list else False)),
        'find_what' : ((type(find_what) is dict) & (False not in [type(x) is str for x in find_what.keys() if type(find_what) is dict]) & (False not in [type(x) is str for x in find_what.values() if type(find_what) is dict])) | \
                      ((type(find_what) is list) & (False not in [type(x) is str for x in find_what if type(find_what) is list])),
        'sec_table_sync' : type(sec_table_sync) is bool,
    }

    for parm in dict_parmTypeJudge.keys():
        if dict_parmTypeJudge[parm] == False:
            raise parmTypeError(parm)

    # 判断参数connect_condition是否为一个列表，如果为列表，则将列表的0号元素与1号元素分别赋值给mainConnectField、secConnectField
    if connect_condition is list:
        # 主数据框查找条件字段名
        mainConnectField = connect_condition[0]
        # 副数据框查找条件字段名
        secConnectField = connect_condition[1]
    else:
        mainConnectField = connect_condition
        secConnectField = connect_condition

    # 通过查找条件找出内容不同的字段
    def finddfDiffContent():

        # 初始化存储内容不同字段的字典
        dict_DiffContent = {}

        # 遍历主数据框中查询条件的元素
        for condition_content in df[0][mainConnectField].values.tolist():
            # 初始化存储内容不同字段(key)及其对应主副数据框元素(value)的字典
            dict_DiffContentColumns = {}

            # 判断当前查询条件元素在主副数据框中是否唯一
            if (df[0][mainConnectField].values.tolist().count(condition_content) > 1) | (df[1][secConnectField].values.tolist().count(condition_content) > 1):
                # 如果不唯一，则在返回结果的字典中标记并不再查找此查询条件元素的字段内容
                dict_DiffContent[mainConnectField] = 'main or secondary table non unique'
                continue

            # 遍历需要查询的内容不同字段的字段名称
            for main_column in find_what.keys():

                # 确定主数据框当前字段对应的值
                main_DfContent = df[0][df[0][mainConnectField] == condition_content][main_column]
                # 判断字段对应的值是否为空
                if len(main_DfContent.values) == 1:
                    main_DfContent = main_DfContent.values[0]
                elif len(main_DfContent) == 0:
                    main_DfContent = None
                # 确定副数据框当前字段对应的值
                sec_DfContent = df[1][df[1][secConnectField] == condition_content][find_what[main_column]].values
                # 判断字段对应的值是否为空
                if len(sec_DfContent) == 1:
                    sec_DfContent = sec_DfContent[0]
                elif len(sec_DfContent) == 0:
                    sec_DfContent = None

                # 判断当前字段主副数据框的值是否相等
                if main_DfContent != sec_DfContent:
                    # 将不相等的字段及其值存储到dict_DiffContentColumns字典中
                    dict_DiffContentColumns[main_column] = [main_DfContent, sec_DfContent]

            # 判断当前查询条件元素是否存在内容不同的字段
            if (dict_DiffContentColumns.__len__() > 0 if type(dict_DiffContentColumns) is dict else type(dict_DiffContentColumns) is str):
                # 将内容不同的查询条件元素及其比较情况存储到dict_DiffContent字典中
                dict_DiffContent[condition_content] = dict_DiffContentColumns

        return dict_DiffContent

    # 更改主数据框中与副数据框比较字段内容不同值为副数据框的值
    def changeDiffContent():
        # 遍历存在内容不同字段的查找条件元素
        for condition_content in finddfDiffContent().keys():

            # 确定其在主数据框的行索引
            mainDfIndex = df[0][df[0][mainConnectField] == condition_content].index
            if len(mainDfIndex) > 0:
                mainDfIndex = mainDfIndex[0]

            # 判断当前元素是否不存在内容不同的字段，或者其是否为字符串
            if (finddfDiffContent()[condition_content].__len__() == 0) | (type(finddfDiffContent()[condition_content]) is str):
                continue

            # 遍历内容不同的字段名
            for column_name in finddfDiffContent()[condition_content].keys():

                # 更改主数据框中的元素
                df[0].loc[mainDfIndex, column_name] = finddfDiffContent()[condition_content][column_name][1]

        return df[0]

    if sec_table_sync is True:

        return finddfDiffContent(), changeDiffContent()

    elif sec_table_sync is False:

        return finddfDiffContent()

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
