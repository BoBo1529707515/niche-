from selenium import webdriver
import time
from selenium.webdriver.common.by import By
import pandas as pd
# 新建一个空的列表用来储存获取的数据
link_all = []
# 设置浏览器内核路径
# 设置 Edge 浏览器驱动程序路径
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options

# 设置 Edge 浏览器驱动程序路径
edge_driver_path = r'F:\edgedriver_win32 (1)\msedgedriver.exe'

# 创建 EdgeOptions 对象
edge_options = Options()

# 创建 Edge 服务对象
service = Service(edge_driver_path)

# 创建 Edge 浏览器对象
driver = webdriver.Edge(service=service, options=edge_options)
# 获取网页
driver.get("https://www.niche.com/k12/search/best-schools/s/texas/")
# 通过设置sleep时间来控制爬虫的速度，根据情况，也可不用。

rank_all = []
name_all = []
href_all = []
def get_text():
    list_rank = driver.find_elements(By.XPATH, '//*[@id="maincontent"]/main/div[1]/div[3]/section/ol//li/div/div/a/div[2]/div[1]')
    for li_rank in list_rank:
        rank_text = li_rank.text
        # 如果 rank_text 中包含 '#' 字符
        if '#' in rank_text:
            # 找到 '#' 字符在字符串中的位置
            hash_index = rank_text.index('#')
            # 获取 '#' 字符后面二位的信息
            next_char = rank_text[hash_index + 1:hash_index + 3]
        rank_all.append(next_char)
    list_name = driver.find_elements(By.XPATH, '//*[@id="maincontent"]/main/div[1]/div[3]/section/ol/li/div/div/a/div[2]/div[2]/h2')
    for li_name in list_name:
        name = li_name.text

        # 保存数据到列表
        name_all.append(name)
# href  大学详情链接地址
    list_href = driver.find_elements(By.XPATH, '//*[@id="maincontent"]/main/div[1]/div[3]/section/ol/li/div/div/a')
    for li_href in list_href:
        href = li_href.get_attribute('href')

        # 保存数据到列表

        href_all.append(href)

get_text()
# 滚动滑轮下滑， = 2500 为向下滑动的距离，这个数值可以根据实际情况调整。
element = driver.find_element(By.XPATH,'//*[@id="maincontent"]/main/div[1]/div[3]/div[2]')
# 通过设置sleep时间来控制爬虫的速度，根据情况，也可不用。看心情
## 控制鼠标点击进行翻页
driver.execute_script("arguments[0].scrollIntoView();", element)
time.sleep(4)
driver.find_element(By.XPATH,'//*[@id="maincontent"]/main/div[1]/div[3]/div[3]/nav/ul/li[7]').click()

# 通过设置sleep时间来控制爬虫的速度，根据情况，也可不用。看心情
time.sleep(1)

get_text()
# 从第三页开始 循环爬取#######
# 此处page设置需要的页数
page = 4
i = 1

# 由于前面已经获取了两页数据，用page-2 来把页数统一
while i <= page - 2:
    element = driver.find_element(By.XPATH, '//*[@id="maincontent"]/main/div[1]/div[3]/div[2]')
    # 通过设置sleep时间来控制爬虫的速度，根据情况，也可不用。看心情
    ## 控制鼠标点击进行翻页
    driver.execute_script("arguments[0].scrollIntoView();", element)

    # 进行翻页
    driver.find_element(By.XPATH, '//*[@id="maincontent"]/main/div[1]/div[3]/div[3]/nav/ul/li[7]').click()
    # 获取第三页数据
    get_text()
    time.sleep(0.5)
    i += 1
print(rank_all)
print(name_all)
print(href_all)
data = pd.DataFrame({
    '排名':rank_all,
    '名称':name_all,
    '详情链接':href_all,


})

# 保存为 Excel 文件
data.to_excel(r'C:\Users\15297\Desktop\niche初步2.xlsx', index=False)