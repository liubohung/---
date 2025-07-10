from bs4 import BeautifulSoup
import re
import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


# 模擬 HTML 字串
def get_sub_page(html_item: str, sub_root_url: str) -> (str, list):
    # 解析 HTML
    soup = BeautifulSoup(html_item, "html.parser")

    # 選擇 input 節點並提取 onclick 屬性值
    onclick_attr = soup.select_one("input")["onclick"]

    # 刪除不必要的字串並分割為參數
    post_parameters = re.sub(r'document\.fm\.|\.value|"', "", onclick_attr)
    post_parameters = re.sub(r"openWindow\(.*?\);", "", post_parameters).split(";")

    # 建立重大資訊網址連結
    iPostParameters = post_parameters  # 假設選取了特定 index，例如 index = 0

    sub_url = (
        sub_root_url
        + f"{iPostParameters[4]}&{iPostParameters[5]}&off=1&firstin=1&{iPostParameters[3]}&year=2018&month=6&"
        + f"{iPostParameters[2]}&{iPostParameters[1]}&{iPostParameters[0]}&b_date=1&e_date=1&t51sb10=t51sb10"
    )
    # print(sub_url)

    sub_url_response = requests.get(sub_url, verify=False)
    sub_url_response_content = sub_url_response.content.decode("UTF-8")

    subpage_detail = []
    # 檢查請求是否成功
    if sub_url_response.status_code == 200:
        # 解析 HTML
        soup = BeautifulSoup(sub_url_response_content, "html.parser")

        # 有無被ban 查詢
        if "查詢過於頻繁" in soup.text:
            return "查詢過於頻繁", [sub_url]

        # 先確認是否為下市公司
        # 找到 h3
        delist_h3 = soup.find("h3")
        if delist_h3 is not None:
            if "已下市" in delist_h3.text:
                return "已下市", []

        # 根據網頁結構提取所需資訊
        sub_table = soup.find("table", {"class": "hasBorder"})

        for row in sub_table.find_all("tr"):
            # print(row)
            cols = row.find_all("td")
            cols_data = [
                col.text.strip().replace("\n", "").replace("\r", "").replace("\xa0", "")
                for col in cols
            ]
            subpage_detail.append(cols_data)
    else:
        print(f"請求失敗，狀態碼：{sub_url_response.status_code}")
        return "get page error " + sub_url_response.status_code, []
    return "succeed", subpage_detail


if __name__ == "__main__":
    html = """<input type="button" value="詳細資料" onclick="document.fm.seq_no.value=&quot;3&quot;;document.fm.spoke_time.value=&quot;220027&quot;;document.fm.spoke_date.value=&quot;20211107&quot;;document.fm.i.value=&quot;38&quot;;document.fm.co_id.value=&quot;3682&quot;;document.fm.TYPEK.value=&quot;sii&quot;;openWindow(document.fm ,&quot;&quot;);">"""
    sub_root_url = "https://mops.twse.com.tw/mops/web/ajax_t05st01?step=2&colorchg=1&"
    print(get_sub_page(html, sub_root_url))
