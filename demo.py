import requests
import re
import time
import os,sys
import xlwt

def getDataFromTXT():
    rootdir = os.getcwd()
    rootdir += '\input'
    #print(rootdir)
    list = os.listdir(rootdir) #列出文件夹下所有的目录与文件
    put_datas = []
    for i in range(0,len(list)):
        path = os.path.join(rootdir,list[i])
        #getFileName
        Name = ""
        for j in range(0,len(list[i])):
            if list[i][j] == '.':
                break
            Name += list[i][j]
        #print(Name)

        if os.path.isfile(path):
            f = open(path, 'r', encoding='UTF-8')
            str_value = '{"strain_name":"' + Name + '","ssurrn_seq":"' + f.read() + '"}'
            #print(f.read())
            put_data = {
                'jsonStr': str_value
            }
            #print(put_data)
            put_datas.append(put_data)
            f.close()
    return put_datas

#----------------------------replace have (.*?)_s name------------------------
def replace_s(pos, sheet, htm):
    results = re.findall("\"n_mismatch\":(.*?)}", htm, re.S)
    print(len(results))
    cnt = 1
    for result in results:
        print(result)
        print('id = ',cnt, result[0])
        taxon_name = re.search("\"taxon_name\":\"(.*?)\"",result, re.S)
        if taxon_name != None and re.search("_s", taxon_name.group(1), re.S) == None:
            print('find first taxon_name', taxon_name.group(1))
            sheet.write(pos, 1, taxon_name.group(1))
            strain_name = re.search("\"strain_name\":\"(.*?)\"",result, re.S)
            if strain_name != None:
                sheet.write(pos, 2, strain_name.group(1)+'(T)')
            similarity = re.search("\"similarity\":(.*?),",result, re.S)
            if similarity != None:
                sheet.write(pos, 3, similarity.group(1))
            taxonomy = re.search("\"taxonomy\":\"(.*?)\"",result, re.S)
            if taxonomy != None:
                sheet.write(pos, 4, taxonomy.group(1))
            return 'replace successful !!!!!!!!!!!'
            break
        cnt += 1
    return 'replace faile !!!!!!!!!!!!'

def login(url,useName,password):
    #--------------------------login------------------------
    login_data = {
        'txtID': useName,
        'txtPWD': password,
    }
    headers_base = {
        'Accept': '*/*',
        'Accept-Encoding': 'gzip, deflate, br',
        'Accept-Language': 'zh-CN',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.140 Safari/537.36 Edge/17.17134',
        'X-Requested-With': 'XMLHttpRequest',
    }
    session = requests.session()
    content = session.post(url, headers = headers_base, data = login_data)
    print(content.text)

    #--------------------------input------------------------
    # put_data = {
    #     'jsonStr': '{"strain_name":"CTD-04-D1-6011m","ssurrn_seq":"CTTCGACTACCGTGGTCGCCTGCCTCCTTGCGGTCAGCGCAGCGCCTTCGGGTAGAACCAACTCCCATGGTGTGACGGGCGGTGTGTACAAGGCCCGGGAACGTATTCACCGCGGCATGCTGATCCGCGATTACTAGCGATTCCAACTTCATGCCCTCGAGTTGCAGAGGACAATCCGAACTGAGACGACTTTTAAGGATTAACCCTCTGTAGTCGCCATTGTAGCACGTGTGTAGCCCACCCTGTAAGGGCCATGAGGACTTGACGTCATCCCCACCTTCCTCCGGCTTAGCACCGGCAGTCCCATTAGAGTTCCCAACTGAATGATGGCAACTAATGGCGAGGGTTGCGCTCGTTGCGGGACTTAACCCAACATCTCACGACACGAGCTGACGACAGCCATGCAGCACCTGTGTCCCAGTCTCCGAAGAGAAAGCCACATCTCTGTGGCGGTCCGGGCATGTCAAAAGGTGGTAAGGTTCTGCGCGTTGCTTCGAATTAAACCACATGCTCCACCGCTTGTGCGGGCCCCCGTCAATTCCTTTGAGTTTTAATCTTGCGACCGTACTCCCCAGGCGGATTGCTTAATGCGTTAGCTGCGTCACCGAAATGCATGCATCCCGACAACTAGCAATCATCGTTTACGGCGTGGACTACCAGGGTATCTAATCCTGTTTGCTCCCCACGCTTTCGAGCCTCAGCGTCAGTAATGAGCCAGTATGTCGCCTTCGCCACTGGTGTTCTTCCGAATATCTACGAATTTCACCTCTACACTCGGAGTTCCACATACCTCTCTCACACTCAAGACACCCAGTATCAAAGGCAATTCCGAGGTTGAGCCCCGGGATTTCACCCCTGACTTAAATGTCCGCCTACGCTCCCTTTACGCCCAGTAATTCCGAGCAACGCTAGCCCCCTTCGTATTACCGCGGCTGCTGGCACGAAGTTAGCCGGGGCTTCTTCTCCGGGTACCGTCATTATCGTCCCCGGTGAAAGAATTTTACAATCCTAAGACCTTCATCATTCACGCGGCATGGCTGCGTCAGGCTTTCGCCCATTGCGCAAGATTCCCCACTGCTGCCTCCCGTAGGAGTTTGGGCCGTGTCTCAGTCCCAATGTGGCTGATCATCCTCTCAGACCAGCTACTGATCGTCGCCTTGGTGAGCCTTTACCTCACCAACTAGCTAATCAGACGCGGGCCGCTCTAAAGGCGATAAATCTTTCCCCCGAAGGGCACATTCGGTATTAGCACAAGTTTCCCTGAGTTATTCCGAACCTAAAGGCACGTTCCCACGTGTTACTCACCCGTCCGCCACTAAGTCCGAAGACTTCGTTCGACTGCAGGTAGTCCGACGCACG"}'
    # }

    put_datas = getDataFromTXT()
    myjobs = []
    for i in range(0,len(put_datas)):
        s = session.post("https://www.ezbiocloud.net/cl16s/submit_identify_data", data = put_datas[i], verify = False)
        print(s.json())
        tmp = s.json().get('sge_job_id',-1)
        if tmp != -1: 
            myjobs.append(tmp) 

    #--------------------------getans------------------------

    time.sleep(1)
    book = xlwt.Workbook(encoding='utf-8', style_compression=0)
    sheet = book.add_sheet('test', cell_overwrite_ok=True)
    sheet.write(0, 0, 'Name')
    sheet.write(0, 1, 'Top-hit taxon')
    sheet.write(0, 2, 'Top-hit strain')
    sheet.write(0, 3, 'Similarity(%)')
    sheet.write(0, 4, 'Top-hit taxonomy')
    sheet.write(0, 5, 'Length')
    rootdir = os.getcwd()
    rootdir += '\output\\'

    for i in range(0,len(myjobs)):
        getID = {
            'jobs': myjobs[i],
        }
        s = session.get("https://www.ezbiocloud.net/cl16s/poll_job_status_multi", params = getID, verify = False)
        print(s.json())
        cnt = 10
        while cnt >= 0 and s.json().get('jobs')[0].get('status') != 'done':
            cnt -= 1
            time.sleep(10)
            s = session.get("https://www.ezbiocloud.net/cl16s/poll_job_status_multi", params = getID, verify = False)
            print(s.json())

        jobs = s.json().get('jobs')
        if jobs[0].get('status') == 'done':
            doneData = jobs[0].get('doneData')
            sheet.write(i+1, 0, doneData.get('strain_name','None'))
            sheet.write(i+1, 1, doneData.get('result_taxon','None'))
            sheet.write(i+1, 2, doneData.get('result_strain','None')+'(T)')
            sheet.write(i+1, 3, doneData.get('result_similarity','None'))
            sheet.write(i+1, 4, doneData.get('result_taxonomy','None'))
            sheet.write(i+1, 5, doneData.get('strain_length','None'))
            if re.search("_s", doneData.get('result_taxon','None'), re.S) != None:
                print('find *_s +++++++++++++++ replace',doneData.get('result_taxon','None'))
                getID = {
                    'id': doneData.get('strain_uid'),
                }
                s = session.get("https://www.ezbiocloud.net/identify/result", params = getID, verify = False)
                print(replace_s(i+1, sheet, s.text))
    NowTime = time.time()
    str(NowTime)
    book.save(rootdir+str(NowTime)+'.xls')

#---------------get Account-----------------
Account = "RDJ5231@163.com"
Password = "5231RDJ"
f = open('account.txt', 'r', encoding='UTF-8')
Account_ifm = f.read()
print(Account_ifm)
File_Account = re.search("\"Account\":\"(.*?)\"", Account_ifm, re.S)
File_Password = re.search("\"Password\":\"(.*?)\"", Account_ifm, re.S)
if File_Account != None and File_Password != None:
    Account = File_Account.group(1)
    Password = File_Password.group(1)

print(Account, Password)
myurl = "https://www.ezbiocloud.net/loginNew"
login(myurl, Account, Password)
