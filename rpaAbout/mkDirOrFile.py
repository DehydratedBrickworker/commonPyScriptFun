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

mkDirOrFile(path=r'C:\Users\No_13\Desktop\a.wmv', isrecreate=True, mkwhat='111')
