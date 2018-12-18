import requests
import re
import time
import os,sys
import xlwt

Answer = []             #中间答案
WaitQueque = {}         #查询结果set

# 输出最后的结果到Excl
def outputAnswerToExcl() :
    book = xlwt.Workbook(encoding='utf-8', style_compression=0)
    sheet = book.add_sheet('sheet1', cell_overwrite_ok=True)
    # green color
    patternGreen = xlwt.Pattern()
    patternGreen.pattern = xlwt.Pattern.SOLID_PATTERN
    patternGreen.pattern_fore_colour = 3
    # 字体   16代表字体大小1个单位
    font = xlwt.Font()
    font.name = '黑体'
    font.bold = True
    font.height = 0x0118
    # 居中
    al = xlwt.Alignment()
    al.horz = 0x02      # 设置水平居中
    al.vert = 0x01      # 设置垂直居中

    style = xlwt.XFStyle() # Create the Pattern
    style.pattern = patternGreen # Add Pattern to Style
    style.font = font
    style.alignment = al

    sheet.write(0, 0, '英文名称', style)
    sheet.write(0, 1, "备注", style)
    sheet.write(0, 2, "选了", style)
    sheet.write(0, 3, "WGS/chromosome", style)
    sheet.write(0, 4, "Cluster", style)
    sheet.write(0, 5, "Type", style)
    sheet.write(0, 6, "From", style)
    sheet.write(0, 7, "To", style)
    sheet.write(0, 8, "Most similar known cluster", style)
    sheet.write(0, 9, "MIBiG BGC-ID", style)
    # 合并
    patternDarkGray = xlwt.Pattern()
    patternDarkGray.pattern = xlwt.Pattern.SOLID_PATTERN
    patternDarkGray.pattern_fore_colour = 23
    style.pattern = patternDarkGray
    sheet.write_merge(1, 1, 0, 10, '', style)

    style1 = xlwt.XFStyle()
    font.name = '宋体'
    font.height = 0x0118
    style1.font = font
    style1.alignment = al
    pos = 2
    for it in Answer :
        if len(it) == 2 :
            sheet.write(pos, 0, it[0], style1)
            sheet.write(pos, 1, it[1], style1)
            sheet.write_merge(pos+1, pos+1, 0, 10, '', style)
            pos = pos+2
            continue
        # sheet.write(pos, 0, it[0], style1)
        # sheet.write(pos, 1, it[2], style1)
        # sheet.write(pos, 2, it[1], style1)
        posS = pos
        flag = False
        for itt in it[3] :
            waitAns = WaitQueque[itt[1]]
            if len(waitAns) < 2:
                flag = True
                continue
            waitAnsUrl = waitAns[0]
            waitAnsData = waitAns[1]
            if waitAns[1] == [] :
                sheet.write_merge(pos, pos, 4, 9, 'No secondary metabolite clusters were found in the input sequence(s)', style1)
                pos = pos+1
                continue
            posSta = pos
            for ittt in waitAnsData :
                link = 'HYPERLINK("%s";"%s")' % (waitAnsUrl, ittt[0])
                sheet.write(pos, 4, xlwt.Formula(link), style1)
                sheet.write(pos, 5, ittt[1], style1)
                sheet.write(pos, 6, ittt[2], style1)
                sheet.write(pos, 7, ittt[3], style1)
                sheet.write(pos, 8, ittt[4], style1)
                sheet.write(pos, 9, ittt[5], style1)
                pos = pos+1
            if pos == posSta :
                 pos = pos+1
            sheet.write_merge(posSta, pos-1, 3, 3, itt[0], style1)
        if pos <= posS :
            pos = pos+1
        sheet.write_merge(posS, pos-1, 0, 0, it[0], style1)
        if flag == False :
            sheet.write_merge(posS, pos-1, 1, 1, it[2], style1)      
        else :
            sheet.write_merge(posS, pos-1, 1, 1, '获取结果失败，请重新查询这个数据', style1)  
        sheet.write_merge(posS, pos-1, 2, 2, it[1], style1) 
        sheet.write_merge(pos, pos, 0, 10, '', style)
        pos = pos+1
        
    NowTime = time.time()
    book.save('output2\\'+str(NowTime)+'.xls')

# 检查查询的串答案是否准备好
def checkDataIsReady(id) :
    while True : 
        try:
            response = requests.get('https://antismash.secondarymetabolites.org/api/v1.0/status/' + id, timeout=5)    
            break
        except requests.exceptions.ConnectionError:
            print("ConnectionError,retrying!!!!!!!!!!!")
        except requests.exceptions.ConnectTimeout:
            print("connectTimeout,retrying!!!!!!!!!!!") 
        except requests.exceptions.ReadTimeout:
            print("ReadTimeout,retrying!!!!!!!!!!!")
    # print(response.json())
    print(id, response.json()['short_status'])
    if response.json()['short_status'] == 'done':
        return response.json()['result_url']
    elif response.json()['short_status'] == 'failed' :
        return 'failed'
    else :
        return 'False'

# 从最后的htm中获取最后的答案
def getFinalData(url) :
    while True : 
        try:
            response = requests.get(url, timeout=5)    
            break
        except requests.exceptions.ConnectionError:
            print("ConnectionError,retrying!!!!!!!!!!!")
        except requests.exceptions.ConnectTimeout:
            print("connectTimeout,retrying!!!!!!!!!!!") 
        except requests.exceptions.ReadTimeout:
            print("ReadTimeout,retrying!!!!!!!!!!!")
    res = re.findall("<tr(.*?)><td class=\"clbutton (.*?)\"><a href=\"#cluster-\w{1,5}\">(.*?)</a></td><td><a href=\"https://docs.antismash.secondarymetabolites.org/glossary/#(.*?)\" target=\"_blank\">(.*?)</a></td><td class=\"digits\">(.*?)</td><td class=\"digits\">(.*?)</td><td>(.*?)</td><td>(.*?)</td></tr>", response.text)
    result = []
    for it in res :
        # print(it)
        tt = re.search('>(.*?)</a>',it[8])
        if tt == None :
            tt = '-'
        else :
            tt = tt.group(1)
        # print(it[2], it[3], it[5], it[6], it[7], tt)
        result.append([it[2], it[3], it[5], it[6], it[7], tt])
    return result

# 向antismash.secondarymetabolites.org上传查询数据
def submitJobToAntismash(sstr) :
    Header = {"Content-type" : "multipart/form-data; boundary=----WebKitFormBoundaryi5674LecZf5ANJew"}
    data= '''------WebKitFormBoundaryi5674LecZf5ANJew
Content-Disposition: form-data; name="knownclusterblast"

true
------WebKitFormBoundaryi5674LecZf5ANJew
Content-Disposition: form-data; name="clusterblast"

true
------WebKitFormBoundaryi5674LecZf5ANJew
Content-Disposition: form-data; name="subclusterblast"

true
------WebKitFormBoundaryi5674LecZf5ANJew
Content-Disposition: form-data; name="smcogs"

true
------WebKitFormBoundaryi5674LecZf5ANJew
Content-Disposition: form-data; name="asf"

true
------WebKitFormBoundaryi5674LecZf5ANJew
Content-Disposition: form-data; name="tta"

true
------WebKitFormBoundaryi5674LecZf5ANJew
Content-Disposition: form-data; name="fullhmmer"

true
------WebKitFormBoundaryi5674LecZf5ANJew
Content-Disposition: form-data; name="ncbi"

''' + sstr + '''
------WebKitFormBoundaryi5674LecZf5ANJew
Content-Disposition: form-data; name="inclusive"

true
------WebKitFormBoundaryi5674LecZf5ANJew
Content-Disposition: form-data; name="cf_threshold"

0.6
------WebKitFormBoundaryi5674LecZf5ANJew
Content-Disposition: form-data; name="borderpredict"

true
------WebKitFormBoundaryi5674LecZf5ANJew--'''
    res = requests.post("https://antismash.secondarymetabolites.org/api/v1.0/submit",data = data,headers = Header)
    print(res.json())
    return res.json()['id']

# 从结果Html提取  Organism/Name、Strain和 Replicons(WGS)
def getRepliconsFromHtm(sname, name, htm, ifm) :
    tmp = htm.split('\n')
    ans = []
    for i in range(1,len(tmp)-1):
        if re.search("<tr class=", tmp[i],re.S) != None :
            # print(tmp[i])
            find1 = re.findall("target=\"_blank\">(.*?)</a></td>", tmp[i], re.S)
            find2 = re.search("/genomes/static/(\w*?).gif", tmp[i], re.S)
            find5 = re.findall("href=\"/nuccore/(.*?)\">(.*?)</a></td>", tmp[i], re.S)
            # print("Strain : ",find1[1],",Level : ",find2.group(1))
            if find1 != [] :
                Strain = find1[1]
            if find2 != None :
                Level = find2.group(1)
                if Level == "complete" :
                    Level = 4
                elif Level == "threequarters" :
                    Level = 3
                elif Level == "half" :
                    Level = 2
                else :
                    Level = 1
            Replicons = []
            for j in range(0,len(find5)):
                Replicons.append([find5[j][0]]) 

        else :
            find3 = re.findall(">((\w|-)+?)<", tmp[i], re.S) 
            find4 = re.search("\w{4}/\w{2}/\w{2}", tmp[i], re.S)
            # if find3 != [] and find4 != None :
            #     print("Strain : ", Strain, ",Level : ",Level, ", Scaffolds : ",find3[1][0],", Date : ",find4.group(), ", Replicons : ",Replicons)
            if find3 != [] :
                Scaffolds = find3[1][0]
                if Scaffolds == '-' :
                    Scaffolds = -1
            if find4 != None :
                Date = find4.group()
            tt = [Strain,Level,Scaffolds,Date,Replicons,i]
            if ans == [] :
                ans = tt
            else :
                if tt[1] == ans[1] and tt[2] == ans[2] and tt[3] > ans[3] :
                    ans = tt
                elif tt[1] == ans[1] and tt[2] < ans[2] :
                    ans = tt
                elif tt[1] > ans[1] :
                    ans = tt
    if ans[4] == [] :
        # print(tmp[ans[5]])
        find6 = re.search("/Traces/wgs/\?val=(.*?)\"", tmp[ans[5]], re.S)
        if find6 != None:
            print(find6)
            while True : 
                try:
                    print("https://www.ncbi.nlm.nih.gov/Traces/wgs/?val="+find6.group(1))
                    response = requests.get("https://www.ncbi.nlm.nih.gov/Traces/wgs/?val="+find6.group(1), timeout=5)    
                    break
                except requests.exceptions.ConnectionError:
                    print("ConnectionError,retrying!!!!!!!!!!!")
                except requests.exceptions.ConnectTimeout:
                    print("connectTimeout,retrying!!!!!!!!!!!") 
                except requests.exceptions.ReadTimeout:
                    print("ReadTimeout,retrying!!!!!!!!!!!")
            find7 = re.search("href=\"https://www.ncbi.nlm.nih.gov/nuccore/(.*?)\"", response.text, re.S)
            if find7 != None :
                ans[4].append([find7.group(1)])
    print(ans)
    Answer.append([sname, name+ans[0], ifm+'overview 有多个', ans[4]])

# 从搜索结果页面获取 (数据/含表单的Htm)
def getHtmBySearchPage(sname, name, page, ifm) :
    find1 = re.search("<a class=\"page_nav\" href=\"/genome/genomes/(.*?)\?\"", page, re.S)
    # overview只有一个
    if find1 == None :
        find2 = re.search("INSDC: <a href=\"(.*?)\">(.*?)</a></td></tr><tr>", page, re.S)
        find3 = re.search("<tr><td style=\"width:50%\"><a href=\"/nuccore/(.*?)\">"+name+"(.*?), whole genome shotgun sequence", page, re.IGNORECASE)
        if find2 != None and find3 != None : 
            ans = [sname, name+find3.group(2), ifm+'overview 只有一个', [[find2.group(2)]]]
        else :
            ans = [sname, 'error']
        print(ans)
        Answer.append(ans)
    # overview有多个
    else :
        print(find1.group(1))
        url2 = "https://www.ncbi.nlm.nih.gov/genomes/Genome2BE/genome2srv.cgi?action=GetGenomes4Grid&genome_id="+find1.group(1)+"&genome_assembly_id=&king=Bacteria&mode=2&flags=1&page=1&pageSize=100"
        print(url2)
        while True : 
            try:
                response2 = requests.get(url2, timeout=5) 
                break
            except requests.exceptions.ConnectionError:
                print("ConnectionError,retrying!!!!!!!!!!!")
            except requests.exceptions.ConnectTimeout:
                print("connectTimeout,retrying!!!!!!!!!!!") 
            except requests.exceptions.ReadTimeout:
                print("ReadTimeout,retrying!!!!!!!!!!!")
        # print(response2.text)
        getRepliconsFromHtm(sname, name, response2.text, ifm)

# 搜索属
def getFromNcbiUseGenus(sname, url, ifm) :
    while True : 
        try:
            response = requests.get(url, timeout=5)    
            break
        except requests.exceptions.ConnectionError:
            print("ConnectionError,retrying!!!!!!!!!!!")
        except requests.exceptions.ConnectTimeout:
            print("connectTimeout,retrying!!!!!!!!!!!") 
        except requests.exceptions.ReadTimeout:
            print("ReadTimeout,retrying!!!!!!!!!!!")
    # 找不到属
    if re.search("The following term was not found in Genome", response.text, re.S) != None :
        print("NO FIND!!!!!!!!!!!!!!!!!!")
        Answer.append([sname, '找不到属和种'])
    # 找的到属
    else :
        find = re.search("link_uid=(.*?)\"><b>(.*?)</b>(.*?)</a>", response.text, re.S)
        if find == None :
            getHtmBySearchPage(sname, sname, response.text, '找的属 ')
            return 
        print(find.group(1),find.group(2),find.group(3))
        if find.group(3) != '' :
            name = find.group(2)+find.group(3)
        else :
            name = find.group(2)+' sp. '
        while True : 
            try:
                print("https://www.ncbi.nlm.nih.gov/genome/genomes/"+find.group(1))
                response1 = requests.get("https://www.ncbi.nlm.nih.gov/genome/"+find.group(1), timeout=5) 
                break
            except requests.exceptions.ConnectionError:
                print("ConnectionError,retrying!!!!!!!!!!!")
            except requests.exceptions.ConnectTimeout:
                print("connectTimeout,retrying!!!!!!!!!!!") 
            except requests.exceptions.ReadTimeout:
                print("ReadTimeout,retrying!!!!!!!!!!!")
        # print(response1.text)
        getHtmBySearchPage(sname, name, response1.text, ifm)

# 搜索属种
def getFromNcbiUseStrain(tmp) :
    url = "https://www.ncbi.nlm.nih.gov/genome/"
    flag = True
    for i in tmp :
        if i == ' ' :
            flag = False
    tmpp = tmp.replace(' ','+')
    url = url+"?term="+tmpp
    print(url)
    while True : 
        try:
            response = requests.get(url, timeout=5)    
            break
        except requests.exceptions.ConnectionError:
            print("ConnectionError,retrying!!!!!!!!!!!")
        except requests.exceptions.ConnectTimeout:
            print("connectTimeout,retrying!!!!!!!!!!!") 
        except requests.exceptions.ReadTimeout:
            print("ReadTimeout,retrying!!!!!!!!!!!")
    # print(response.text)
    # 找不到属种
    if flag == True or re.search("The following term was not found in Genome", response.text, re.S) != None :
        url1 = ""
        for i in range(0,len(url)):
            if url[i] == '+' :
                url1 = url[:i]
                break
        if flag == True :
            url1 = url
        print(url1)
        getFromNcbiUseGenus(tmp, url1, '找的属 ')
    # 找到属种
    else :
        getHtmBySearchPage(tmp, tmp, response.text, '找的种 ')


# Answer = [['Bacillus aryabhattai', 'Bacillus aryabhattai K13', 'fffff', [['CP001879.1', 'bacteria-1b641405-37e8-4359-9e9e-91c1d71349e5'], ['CP001880.1', 'bacteria-c978cce6-1c90-48c5-95e4-3297024e4a92'], ['NZ_CP024037.1', 'bacteria-39c752a6-d4f0-4f48-adf1-90a9e11a871f']]]]

# getFinalData("https://antismash.secondarymetabolites.org/upload/bacteria-1b641405-37e8-4359-9e9e-91c1d71349e5/index.html")
# submitJobToAntismash(ans)
# getFromNcbiUseStrain("Bacillus aryabhattai")

f = open('input2/input.txt', 'r', encoding='UTF-8')
Input = f.read()
f.close()
for it in Input.split('\n') :
    tmp = it.replace('\n','')
    ttmp = tmp.replace(' ','')
    if tmp == '' or ttmp == '':
        continue 
    print("---------------",it,"--------------------")    
    getFromNcbiUseStrain(tmp)
    print()

# 将第一步获取的答案post进行查询
for i in range(0, len(Answer)) :
    print(Answer[i])
    if len(Answer[i]) < 4:
        continue
    if Answer[i][3] == [] :
        Answer[i][2] =  Answer[i][2] + '查不到对应的Replicons(WGS)，跳过'
    for j in range(0, len(Answer[i][3])) :
        jobID = submitJobToAntismash(Answer[i][3][j][0])
        Answer[i][3][j].append(jobID)
        WaitQueque[jobID] = []
        # WaitQueque[Answer[i][3][j][1]] = []
        # print(Answer[i][3])

# 不断的查询每个post的状态，直到每个都输出结果
flag = True
while flag == True :
    # time.sleep(600)
    flag = False
    for key in WaitQueque :
        status = checkDataIsReady(key)
        if status != 'False' :
            if status == 'failed' :
                WaitQueque[key].append('failed')
                continue
            status = 'https://antismash.secondarymetabolites.org/'+status
            WaitQueque[key].append(status)                        # 结果的网址
            WaitQueque[key].append(getFinalData(status))          # 提取后的最终结果
        else :
            flag = True
    print('-----------------------------------')
outputAnswerToExcl()
