import json
import os

import xlwings as xw
import FileParser.Cache as ca

def parseJsonFile (filePath):
    with open(filePath) as f:
        data = json.load(f)
    #print(data)
    return data

SheetNames = ['基本属性', '守门员属性', '精神属性', '身体属性', '技术属性', '隐藏属性', '位置', '球员性格']

AttrLabels = [['日期', '名字', '  CA  ', '  PA  ', '  RCA  ', '身价', '身高', '体重', '世界声望', '当前声望', '本国声望', '能力评分', '潜力评分']
    , ['日期', '制空能力', '拦截传中', '指挥防守', '神经指数', '手控球', '大脚开球', '一对一', '反应', '出击倾向', '击球倾向', '手抛球']
    , ['日期', '侵略性', '预判', '勇敢', '镇定', '集中', '决断', '意志力', '想象力', '领导力', '无球跑动', '防守站位', '团队合作', '视野', '工作投入']
    , ['日期', '爆发', '灵活', '平衡', '弹跳', '左脚', '体质', '速度', '右脚', '耐力', '强壮']
    , ['日期', '稳定性', '肮脏', '大赛', '受伤', '多面性']
    , ['日期', '角球', '传中', '盘带', '射门', '接球', '任意球', '头球', '远射', '界外球', '盯人', '传球', '点球', '抢断', '技术']
    , ['日期', '门将', '左后卫', '中后卫', '右后卫', '后腰', '左边卫', '右边卫', '左前卫', '中场', '右前卫', '左边锋', '前腰', '右边锋', '前锋']
    , ['日期', '适应性', '野心', '忠诚', '抗压', '职业', '体育精神', '情绪控制', '争论']
]

def writeXlsxHeader(wb):
    wb.sheets[0].name = SheetNames[0]
    for i in range(len(SheetNames)):
        if i == 0:
            wb.sheets[0].name = SheetNames[0]
        else:
            wb.sheets.add(SheetNames[i], None, SheetNames[i - 1])
        sht = wb.sheets(i+1)
        headers = AttrLabels[i]
        for j in range(len(headers)):
            sht[0, j].value = headers[j]

    wb.sheets(1).activate()
    return

def dateExisted(date, sht):
    rng = sht.range('a1').expand('table')
    nrows = rng.rows.count
    existDates = sht.range(f'a1:a{nrows}').value

    # 已经存在则不写入
    if date in existDates:
        return [True, nrows]
    else:
        return [False, nrows]

def writeCommonAttributes(personData, date, name, sht):
    #读取第一列，判断日期是否已存在
    temp = dateExisted(date, sht)
    if temp[0]:
        return

    nrows = temp[1]
    # 不是第一次写入
    if nrows > 2:
        # 清理差值行
        nrows -= 1
        sht[nrows, 0:13].delete('up')
        # 清理之前单元格的颜色
        sht[nrows, 0:13].color = (255, 255, 255)

    #['日期', '名字', 'CA', 'PA', 'RCA', '身价', '身高', '体重', '世界声望', '当前声望', '本国声望', '能力评分', '潜力评分']
    sht[nrows, 0].api.NumberFormat = "@"
    sht[nrows, 0].value = date
    sht[nrows, 1].value = name
    sht[nrows, 2].value = personData['CA']
    sht[nrows, 3].value = personData['PA']
    sht[nrows, 4].value = personData['RCA']
    sht[nrows, 5].value = personData['Value']
    sht[nrows, 6].value = personData['Height']
    sht[nrows, 7].value = personData['Weight']
    sht[nrows, 8].value = personData['WorldReputation']
    sht[nrows, 9].value = personData['CurrentReputation']
    sht[nrows, 10].value = personData['HomeReputation']
    sht[nrows, 11].value = personData['ActualRating']
    sht[nrows, 12].value = personData['PotentialRating']

    sht.autofit()

    # 有2行以上的数据时计算差值
    nrows += 1
    if nrows > 2:
        sht[nrows, 0].value = '成长'
        for col in range(2, 13):
            sht[nrows, col].value = sht[nrows - 1, col].value - sht[1, col].value
            # 填充颜色
            if sht[nrows, col].value > 0:
                sht[nrows, col].color = (105, 229, 137)
            elif sht[nrows, col].value < 0:
                sht[nrows, col].color = (242, 127, 100)
    return

def getAttributeList(personDate, index):
    if index == 1:
        return personDate['GoalKeeperAttributes']
    elif index == 2:
        return personDate['MentalAttributes']
    elif index == 3:
        return personDate['PhysicalAttributes']
    elif index == 4:
        return personDate['HiddenAttributes']
    elif index == 5:
        return personDate['TechnicalAttributes']
    elif index == 6:
        return personDate['Positions']
    elif index == 7:
        return personDate['PersonalityAttributes']

#['基本属性', '守门员属性', '精神属性', '身体属性', '技术属性', '隐藏属性', '位置', '球员性格']
def writeAttributes(personData, date, name, wb):
    for i in range(len(SheetNames)):
        if i == 0:
            writeCommonAttributes(personData, date, name, wb.sheets(1))
        else:
            sht = wb.sheets(i+1)
            temp = dateExisted(date, sht)
            if temp[0]:
                continue
            nrows = temp[1]

            attributes = getAttributeList(personData, i)
            ncols = len(attributes) + 1
            # 不是第一次写入
            if nrows > 2:
                # 清理差值行
                nrows -= 1
                sht[nrows, 0:ncols].delete('up')
                # 清理之前单元格的颜色
                sht[nrows, 0:ncols].color = (255, 255, 255)

            sht[nrows, 0].api.NumberFormat = "@"
            sht[nrows, 0].value = date
            j = 0
            for key, attr in attributes.items():
                j += 1
                sht[nrows, j].value = attr

            # 有2行以上的数据时计算差值
            nrows += 1
            if nrows > 2:
                sht[nrows, 0].value = '成长'
                for col in range(1, ncols):
                    sht[nrows, col].value = sht[nrows - 1, col].value - sht[1, col].value
                    # 填充颜色
                    cell = sht[nrows, col]
                    if cell.value < 0:
                        cell.color = (242, 127, 100)
                    elif cell.value == 0:
                        cell.color = (255, 255, 255)
                    elif 0 < cell.value <= 1:
                        # 浅绿
                        cell.color = (156, 238, 158)
                    elif cell.value <= 3:
                        # 深绿
                        cell.color = (0, 255, 0)
                    elif cell.value <= 5:
                        # 浅蓝
                        cell.color = (105, 209, 230)
                    else:
                        # 深蓝
                        cell.color = (0, 100, 255)

            sht.autofit()
    return

def scanAllFiles(path):
    global wb, app, xlsxName, currentDate
    files = os.listdir(path)

    # 记录上一个处理的人名，避免重复打开xlsx
    lastPerson = ''
    isFirstPerson = True

    wb = None

    #扫描文件夹下所有文件
    for file in files:

        file_d = os.path.join(path, file)

        if os.path.isdir(file_d):
            #TODO
            print(file_d)
        else:
            #跳过非json文件
            if not file.endswith(".json"):
                continue
            if ca.checkInCache(file, path):
                continue

            print("Parse file:" + file_d)
            personData = parseJsonFile(file_d)

            #检查是否有同名xlsx文件，创建xlsx文件的操作句柄
            splitedStr = file.split('-')
            personName = splitedStr[0] + "-" + splitedStr[1]
            currentDate = splitedStr[2].split('.')[0]

            # 处理其他人时保存上一个人的文件
            if (personName != lastPerson):
                if (isFirstPerson != True):
                    wb.save(xlsxName)
                    wb.close()
                    app.quit()

                xlsxName = os.path.join(path, personName + ".xlsx")

                #新建或者打开文件
                if not os.path.exists(xlsxName):
                    print("Create " + xlsxName)
                    app = xw.App(visible=True, add_book=False)
                    wb = app.books.add()
                    #新建文件时写入文件头
                    writeXlsxHeader(wb)
                else:
                    print("Open " + xlsxName)
                    app = xw.App(visible=True, add_book=False)
                    app.display_alerts = False
                    app.screen_updating = False
                    wb = app.books.open(xlsxName)

            #写入属性
            writeAttributes(personData, currentDate, splitedStr[1], wb)
            lastPerson = personName
            isFirstPerson = False

            # 写入缓存
            ca.saveCache(file, path)

    if wb is not None:
        wb.save(xlsxName)
        wb.close()
        app.quit()
