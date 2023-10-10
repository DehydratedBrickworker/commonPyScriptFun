def writeToExcel(path:str, values:dict, sheetname=0, columns:list=None, istranspose=False):
    '''

    :param path: Excel文件路径
    :param values: 写入位置:写入的值 组成的字典
    :param sheetname: 页名
    :param columns: 列名组成的列表，默认值为None->跳过写入列名的步骤
    :param istranspose: 是否转置
    :return:
    '''
    import xlwings as xw

    app = xw.App(visible=False, add_book=False)
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