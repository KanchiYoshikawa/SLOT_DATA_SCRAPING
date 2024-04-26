from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from xlwt import Workbook
import time
import urllib.request
import os
import datetime
import openpyxl
import glob


def Output_excel(sheet, count=0):
    #台数の取得
    element = driver.find_elements_by_xpath("/html/body/div[5]/div[1]/div[2]/table/tbody/tr")
    machine_num = len(element) - 2
    
    #機種名の記入
    element = driver.find_element_by_xpath("/html/body/div[5]/h2/a")
    sheet["A" + str(count + 3)] = str(element.text)

    #データの収集
    for j in range(3, machine_num + 3, 1):

        #台番号
        element = driver.find_element_by_xpath("/html/body/div[5]/div[1]/div[2]/table/tbody/tr[" + str(j-1) + "]/td[1]/span[1]")
        actions = ActionChains(driver)
        actions.move_to_element(element)
        actions.perform()
        machine_No.append(element.text)
        sheet["B" + str(count + j)] = int(element.text)

        element = driver.find_element_by_xpath("/html/body/div[5]/div[1]/div[2]/table/tbody/tr[" + str(j-1) + "]/td[2]")
        total_game.append(element.text)
        try:
            sheet["C" + str(count + j)] = int(element.text)
        except:
            sheet["C" + str(count + j)] = str(element.text)

        element = driver.find_element_by_xpath("/html/body/div[5]/div[1]/div[2]/table/tbody/tr[" + str(j-1) + "]/td[3]")
        BB_num.append(element.text)
        try:
            sheet["D" + str(count + j)] = int(element.text)
        except:
            sheet["D" + str(count + j)] = str(element.text)

        element = driver.find_element_by_xpath("/html/body/div[5]/div[1]/div[2]/table/tbody/tr[" + str(j-1) + "]/td[4]")
        RB_num.append(element.text)
        try:
            sheet["E" + str(count + j)] = int(element.text)
        except:
            sheet["E" + str(count + j)] = str(element.text)

        element = driver.find_element_by_xpath("/html/body/div[5]/div[1]/div[2]/table/tbody/tr[" + str(j-1) + "]/td[7]")
        BB_prob.append(element.text)
        try :
            sheet["F" + str(count + j)] = int(element.text)
        except :
            sheet["F" + str(count + j)] = "--"

        element = driver.find_element_by_xpath("/html/body/div[5]/div[1]/div[2]/table/tbody/tr[" + str(j-1) + "]/td[8]")
        RB_prob.append(element.text)
        try :
            sheet["G" + str(count + j)] = int(element.text)
        except :
            sheet["G" + str(count + j)] = "--"

        element = driver.find_element_by_xpath("/html/body/div[5]/div[1]/div[2]/table/tbody/tr[" + str(j-1) + "]/td[5]")
        total_prob.append(element.text)
        try :
            sheet["H" + str(count + j)] = int(element.text)
        except :
            sheet["H" + str(count + j)] = "--"



URL = "http://fe.site777.tv/data/biglobe/login.php"

#URLのページを開く
options = Options()
options.add_argument('--headless')
driver = webdriver.Chrome('.\chromedriver.exe')#, options=options)
driver.implicitly_wait(40)
driver.get(URL)

#Windowのサイズを最大化する
driver.maximize_window()

#BIGLOBEのログイン
driver.find_element_by_id("BiglobeId").send_keys("siteseven.data@gmail.com")
driver.find_element_by_id("BiglobePwd").send_keys("Crawler777")
#なぜか要素をクリックできなかったので要素の場所を指定してクリックした
element = driver.find_element_by_xpath("/html/body/div[1]/form/div/div/div[1]/input")
loc = element.location
x, y = loc['x'], loc['y']
actions = ActionChains(driver)
actions.move_by_offset(x, y)
actions.click()
actions.perform()

#神奈川のホール一覧に飛ぶリンクをクリック
#element = driver.find_element_by_xpath("/html/body/div[4]/div[2]/div[1]/div[3]/div/div/div[2]/dl[2]/dd")
element = driver.find_element_by_xpath("/html/body/div[4]/div[2]/div[1]/div[3]/div/div/div[2]/dl[2]/dd/a[2]")
actions = ActionChains(driver)
actions.move_to_element(element)
actions.perform()
element.click()


#希望地域の選択
#か行のホール
#/html/body/div[4]/div[2]/div[1]/form/div[1]/div/div/dl[8]/dd/ul/li[1]/input
#/html/body/div[4]/div[2]/div[1]/form/div[1]/div/div/dl[8]/dd/ul/li[4]/input
area_number = [3]
for i in area_number:
    
    driver.find_element_by_xpath("/html/body/div[4]/div[2]/div[1]/form/div[1]/div/div/dl[2]/dd/ul/li[" + str(i) + "]/input").click()    
    #検索ボタンのクリック
    element = driver.find_element_by_xpath("/html/body/div[4]/div[2]/div[1]/form/div[1]/div/div/div[2]/p/a")
    actions = ActionChains(driver)
    actions.move_to_element(element)
    actions.perform()
    driver.find_element_by_xpath("/html/body/div[4]/div[2]/div[1]/form/div[1]/div/div/div[2]/p/a").click()
    
    time.sleep(3)
    
    #検索地域のホール数の検索
    element = driver.find_elements_by_xpath("/html/body/div[4]/div[2]/div[1]/div/form/div")
    Hole_number = len(element)
    
    #全てのホールにアクセスを開始する
    for j in range(1, Hole_number+1, 1):
        #ホールの選択
        element = driver.find_element_by_xpath("/html/body/div[4]/div[2]/div[1]/div/form/div[" + str(j) + "]/div/div/p[1]/a")
        actions = ActionChains(driver)
        actions.move_to_element(element)
        actions.perform()
        driver.find_element_by_xpath("/html/body/div[4]/div[2]/div[1]/div/form/div[" + str(j) + "]/div/div/p[1]/a").click()
        
        #ホール名の出力
        element = driver.find_element_by_xpath("/html/body/div[5]/div/div[2]/div/h1[1]")
        print(element.text)
        
        
        #データの格納先の定義
        machine_No = []
        total_game = []
        BB_num = []
        RB_num = []
        BB_prob = []
        RB_prob = []
        total_prob = []

        #ホール名の取得しフォルダを作成する
        element = driver.find_element_by_xpath("/html/body/div[5]/div/div[2]/div/h1[1]")
        hole_dir = "./data/" + str(element.text)
        os.makedirs(hole_dir, exist_ok=True)

        #データの日付を取得しフォルダを作成する
        dt_now = datetime.datetime.today() - datetime.timedelta(days=1)
        day_dir = hole_dir + "/" + datetime.datetime.strftime(dt_now, '%Y-%m-%d')
        os.mkdir(day_dir)

        excel_name_all = day_dir + "/全機種データ一覧.xlsx"
        wb_all = openpyxl.Workbook()
        sheet_all = wb_all['Sheet']

        sheet_all["B2"] = "台番号"
        sheet_all["C2"] = "総回転数"
        sheet_all["D2"] = "BB回数"
        sheet_all["E2"] = "RB回数"
        sheet_all["F2"] = "BB確率"
        sheet_all["G2"] = "RB確率"
        sheet_all["H2"] = "合算確率"

        #スロットの機種の選択
        try:
            i = 2
            #全機種データのexce1記入のための定義
            count = 0
            while(True):

                #i番目の機種を選択する
                element = driver.find_element_by_xpath("/html/body/form[5]/div/div/center/table/tbody/tr[" + str(i) + "]/td/ul/li[1]/input")
                actions = ActionChains(driver)
                actions.move_to_element(element)
                actions.perform()

                #20スロのデータが存在する場合の処理
                if element.get_attribute("value") == "出玉データ":

                    #i番目の機種のデータをクリックする
                    element.click()

                    #台数の取得
                    element = driver.find_elements_by_xpath("/html/body/div[5]/div[1]/div[2]/table/tbody/tr")
                    machine_num = len(element) - 2

                    #機種名を取得しフォルダを作成
                    element = driver.find_element_by_xpath("/html/body/div[5]/h2/a")
                    #print("機種名　" + str(element.text))
                    type_dir = day_dir + "/" + str(element.text)
                    os.mkdir(type_dir)

                    #全データのexcelのデータ収集
                    #データを取得しexcelへの記入
                    Output_excel(sheet_all, count)
                    #excelファイルの保存
                    wb_all.save(excel_name_all)
                    count += machine_num
                    
                    #グラフのページへ移動
                    driver.find_element_by_xpath("/html/body/div[5]/div[1]/div[1]/div[1]/p/a").click()


                    #グラフデータの保存
                    graph_num = 0
                    try:
                        while(True):
                            for j in range(1, 8, 1):
                                graph_num += 1
                                element = driver.find_element_by_xpath("/html/body/div[5]/div[2]/dl[" + str(j) + "]/dd/a/img")
                                actions = ActionChains(driver)
                                actions.move_to_element(element)
                                actions.perform()
                                url = element.get_attribute("src")

                                element = driver.find_element_by_xpath("/html/body/div[5]/div[2]/dl[" + str(j) + "]/dt/a")
                                daiban = element.text
                                graph_name = type_dir + "/" + str(daiban[3:]) + ".png"
                                urllib.request.urlretrieve(url, graph_name)

                            #次の一覧に移動するボタンがあればクリックしてなければエラーが出るので処理がスキップされる
                            element = driver.find_element_by_xpath("/html/body/div[5]/dl/dd/a").click()

                    except:
                        time.sleep(1)

                    i += 1
                    #ページを機種一覧に戻す
                    element = driver.find_element_by_xpath("/html/body/div[5]/ul/li[1]/a")
                    actions = ActionChains(driver)
                    actions.move_to_element(element)
                    actions.perform()
                    element.click()

                #20スロのデータが存在しない場合の処理
                else:
                    #ホール一覧へ戻る
                    element = driver.find_element_by_xpath("/html/body/div[4]/p/a[4]")
                    actions = ActionChains(driver)
                    actions.move_to_element(element)
                    actions.perform()
                    driver.find_element_by_xpath("/html/body/div[4]/p/a[4]").click()
                    os.rmdir(day_dir)
                    if len(glob.glob(hole_dir + "/*")) == 0:
                        os.rmdir(hole_dir)
                    break
        except:
            if i==2:
                print("20スロのデータがありません")
                os.rmdir(day_dir)
                if len(glob.glob(hole_dir + "/*")) == 0:
                    os.rmdir(hole_dir)
            else:
                print("20スロのデータ収集完了")

            #ホール一覧へ戻る
            element = driver.find_element_by_xpath("/html/body/div[4]/p/a[4]")
            actions = ActionChains(driver)
            actions.move_to_element(element)
            actions.perform()
            driver.find_element_by_xpath("/html/body/div[4]/p/a[4]").click()
    
    element = driver.find_element_by_xpath("/html/body/div[4]/div[1]/p[1]/a[2]")
    actions = ActionChains(driver)
    actions.move_to_element(element)
    actions.perform()
    driver.find_element_by_xpath("/html/body/div[4]/div[1]/p[1]/a[2]").click()
    
    #神奈川のホール一覧に飛ぶリンクをクリック
    #element = driver.find_element_by_xpath("/html/body/div[4]/div[2]/div[1]/div[3]/div/div/div[2]/dl[2]/dd")
    element = driver.find_element_by_xpath("/html/body/div[4]/div[2]/div[1]/div[1]/div/div/div[2]/dl[2]/dd/a[2]")
    actions = ActionChains(driver)
    actions.move_to_element(element)
    actions.perform()
    element.click()


driver.quit()