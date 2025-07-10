# 下載 公開資訊觀測站 總經理異動資料

import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import time
import os
import random
import pandas as pd

from detail_button import get_sub_page

get_url = "https://mopsov.twse.com.tw/mops/web/t51sb10_q1"
sub_root_url = "https://mopsov.twse.com.tw/mops/web/ajax_t05st01?step=2&colorchg=1&"


# 參數說明
# 參數欄位
# url: 網址
# cond_firm_type: 限縮條件 "上市公司" "上櫃公司" "興櫃公司" "公開發行公司"
# year_timing: 時間年度
# month_timing: 時間月份
# PCount: 每月資料
# keyWord: 搜尋關鍵字
# condition2: 第二條件 ["且含", "或含" ,"不含"]
# second_keyWord: 第二條件關鍵字


def getpage(
    url="",
    cond_firm_type="上櫃公司",
    year_timing="110",
    month_timing="全年度",
    PCount="100",
    keyWord="總經理異動",
    condition2="不含",
    second_keyWord="副總經理異動",
):
    if url == "":
        print("url is empty")
        return
    global driver
    options = webdriver.ChromeOptions()
    options.add_experimental_option("detach", True)  # 保持瀏覽器開啟

    driver = webdriver.Chrome(options=options)

    driver.get(url)
    # find name = form1
    cond_table = driver.find_element(By.NAME, "form1")
    # print(cond_table.text)

    data_df = pd.DataFrame()
    try:
        # 等待頁面載入 (可根據實際情況調整)
        time.sleep(2)

        # 找到含有 "請選擇市場別/產業別:" 文字的 <td>
        market_td = driver.find_element(
            By.XPATH, "//td[contains(text(), '請選擇市場別/產業別:')]"
        )

        # 找到該 <td> 的內部 <input> 元素
        market_input = market_td.find_element(By.XPATH, ".//input")

        # 點擊該 <input>
        market_input.click()

        time.sleep(1)
        # 選擇「市場別」為「上市公司」
        market_select = Select(driver.find_element(By.NAME, "KIND"))
        market_select.select_by_visible_text(cond_firm_type)
        time.sleep(1)

        # 輸入關鍵字「總經理異動」
        keyword_input = driver.find_element(By.NAME, "keyWord")
        keyword_input.clear()
        keyword_input.send_keys(keyWord)
        time.sleep(1)

        # 選擇條件2為「不含」
        condition2_select = Select(driver.find_element(By.NAME, "Condition2"))
        condition2_select.select_by_visible_text(condition2)
        time.sleep(1)

        # 在條件2輸入「副總經理異動」
        keyword2_input = driver.find_element(By.NAME, "keyWord2")
        keyword2_input.clear()
        keyword2_input.send_keys(second_keyWord)
        time.sleep(1)

        # 設定年度為113
        year_input = driver.find_element(By.NAME, "year")
        year_input.clear()
        year_input.send_keys(year_timing)
        time.sleep(1)

        # 設定月份為「全年度」
        month_select = Select(driver.find_element(By.NAME, "month1"))
        month_select.select_by_visible_text(month_timing)
        time.sleep(1)

        # 點擊「搜尋」按鈕
        search_button = driver.find_element(
            By.XPATH, "//div[@id='search_bar1']//input[@value=' 搜尋 ']"
        )

        search_button.click()

        # 等待搜尋結果載入
        time.sleep(3)

        # 如果有 font 是 查無所需資料 ，則跳出
        try:
            no_data_font = driver.find_element(
                By.XPATH, "//font[contains(text(), '查無所需資料')]"
            )
            if no_data_font:

                print("查無所需資料")
                return
            # print(no_data_font.text)
            # 如果有查無所需資料，則跳出
            print("查無所需資料")
        except:
            pass
        # # 找到 id 為 table01 的表格
        # # 找 該網頁新載入的表格
        PCount_select = Select(driver.find_element(By.ID, "PCount"))

        PCount_select.select_by_value(PCount)  # 選擇顯示 100 筆

        # 點擊 Go 按鈕
        go_button = driver.find_element(By.ID, "PCountbtn")
        go_button.click()

        # 等待資料重新載入
        time.sleep(3)

        # 找到 form 中的 span
        try:
            form_span = driver.find_element(By.XPATH, "//form/span")
            print(form_span.text)
        except:
            # 如找不到，表示只有一頁
            form_span = None
        if form_span is not None:
            # 找到最後一個 button
            last_button = form_span.find_element(By.XPATH, "//span/button[last()]")
            # print(last_button.text)

            # 找到最後一個 button的 前一個 button
            last_button_pre = form_span.find_element(
                By.XPATH, "//span/button[last()-1]"
            )
            # print(last_button_pre.text)

        # 如果最後一個符號是 >，表示 這個 span 是所有 頁數的按鈕
        # 如果最後一個符號是數字，這個 span 包報含頁數的按鈕

        # 確認點擊多少次 > 按鈕
        if form_span is None:
            print("未找到頁數按鈕_僅一頁")
            click_count = 1
        elif last_button.text == ">":
            # 如果是 >，表示不是最後一頁
            click_count = int(last_button_pre.text)
        else:
            # 如果是數字，表示是最後一頁
            click_count = int(last_button.text)

        print("click_count:", click_count)

        # 根據頁數點擊
        for i in range(click_count):
            # 獲取搜尋結果表格的 HTML
            soup = BeautifulSoup(driver.page_source, "html.parser")
            table = soup.find(
                "table", {"class": "hasBorder"}
            )  # 假設表格類別為 hasBorder

            table_data = []
            sub_page_error_list = []
            # 如果表格存在，解析內容
            if table:
                rows_driver = driver.find_elements(
                    By.XPATH, "//table[@class='hasBorder']/tbody/tr"
                )
                for row in rows_driver:
                    cols = row.find_elements(By.TAG_NAME, "td")
                    # 如果沒有抓到任何 <td>，表示這一列是表格標題
                    if cols == []:
                        continue
                    # 抓取這一列之資料
                    col_data = [col.text.strip() for col in cols]
                    # 找到 input button 的 dom
                    details_button = cols[-1].get_attribute("innerHTML")

                    # 抓取子網頁中之資訊
                    rest_time = 0
                    out_count = 0
                    while True:
                        try:
                            is_sus, sub_page_content = get_sub_page(
                                details_button, sub_root_url
                            )

                            if "查詢過於頻繁" in is_sus:
                                rest_time += 1
                                print(f"查詢過於頻繁<--休息一下 {rest_time} 次")
                                time.sleep(random.uniform(9, 15))
                                continue
                            if rest_time > 5:
                                print("放棄此筆資料額外資料")
                                col_data.append(["查詢過於頻繁", sub_page_content])
                                break
                            # 如果抓取成功
                            if is_sus == "succeed":
                                col_data.pop()
                                col_data.append(sub_page_content)
                                break
                            elif is_sus == "已下市":
                                col_data.pop()
                                col_data.append("已下市無資料")
                                break
                        except Exception as e:
                            # 可能是因為被封，所以抓不到，故等待六秒後再次嘗試
                            out_count += 1
                            if out_count > 5:
                                print("放棄此筆資料額外資料")
                                col_data.append(["查詢過於頻繁", sub_page_content])
                                break
                            print("error:", e)
                            print("陷入休息情況～～")
                            time.sleep(random.uniform(6, 10))

                    # 暫停秒數，避免被封
                    sleep_t = random.uniform(0, 3)
                    time.sleep(sleep_t)
                    print(col_data)

                    # 匯集該筆資料
                    table_data.append(col_data)

            # 將表格資料轉換為 DataFrame
            # 第一列為標題另外用方法去除
            table_s_data = table_data
            table_s_data_df = pd.DataFrame(
                table_s_data,
                columns=[
                    "公司代號",
                    "公司名稱",
                    "異動日期",
                    "序號",
                    "異動事項",
                    "網址",
                ],
            )
            # 合併資料
            data_df = pd.concat([data_df, table_s_data_df])

            # 找到下一頁按鈕
            if click_count == 1:
                break
            else:
                target_button = driver.find_element(
                    By.XPATH, "//span/button[text()='>']"
                )
                target_button.click()

            # 等待資料重新載入
            time.sleep(3)

        data_df.to_csv(
            "./Data/" + year_timing + "_" + cond_firm_type + f"_{keyWord}_change.csv",
            index=False,
            encoding="utf-8-sig",
        )
    except Exception as e:
        print(e)


if __name__ == "__main__":
    # 確定有無 data 資料夾存在 若無則建立
    if not os.path.exists("./Data"):
        os.makedirs("./Data")

    # 設定參數
    # 欲爬取的公司類型
    cond_firm_type_list = ["上市公司", "上櫃公司"]  # "興櫃公司" "公開發行公司"
    # 欲爬取的時間區間
    year_timing = [
        str(i) for i in range(113, 114)
    ]  # range 中為時間區間，為113年到114年
    # 設定月份
    month_timing = "全年度"
    # 頁面顯示筆數
    PCount = "100"
    # 設定關鍵字
    keyWords = [
        "總經理異動",
        "任總經理",
        "總經理變動",
        "總經理聘任",
        "總經理人事",
    ]
    # 第二條件
    condition2 = "不含"  # "且含" "或含" "不含"
    # 第二關鍵字
    second_keyWord = "副總經理"
    # 開始爬蟲
    print("開始執行...")
    for cond_firm_type in cond_firm_type_list:
        for year in year_timing:
            for keyWord in keyWords:
                getpage(
                    get_url,
                    cond_firm_type,
                    year,
                    month_timing,
                    PCount,
                    keyWord,
                    condition2,
                    second_keyWord,
                )
                time.sleep(random.uniform(5, 10))
    # getpage(get_url)
