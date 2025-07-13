import pandas as pd
import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from datetime import datetime


# ============ 【⚙️ 配置信息 你要改成自己的】 ==============
EXCEL_PATH = "Reset_Password.xlsx"

CHROMEDRIVER_PATH = "chromedriver.exe"

WAIT_BETWEEN_ACCOUNTS = 5  # 每个账号之间等待几秒
# ===========================================================


# ✅ ① 读取Excel
print("✅ 读取 Excel 文件...")
df = pd.read_excel(EXCEL_PATH, dtype=str)
print(f"✅ 读取到 {len(df)} 条账号信息。")
print(df)

# ✅ ② 启动浏览器
print("✅ 启动浏览器...")
options = Options()
options.binary_location = "chrome-win64/chrome.exe"
options.add_experimental_option("excludeSwitches", ["enable-logging"])

service = Service(CHROMEDRIVER_PATH)
driver = webdriver.Chrome(service=service, options=options)

#增加TXT日志
BASE_DIR = os.getcwd()
LOG_DIR = os.path.join(BASE_DIR, "logs")
os.makedirs(LOG_DIR, exist_ok=True)
today_str = datetime.now().strftime("%Y%m%d")
log_path = os.path.join(LOG_DIR, f"log_{today_str}.txt")

log_file = open(log_path, "a", encoding="utf-8")
log_file.write("\n==== NEW SESSION ====\n")
log_file.write(f"{'TIMESTAMP':<25}  {'CODICE FISCALE':<25}\tRESULT\n")

# ✅ ③ 遍历每一行账户信息
for idx, row in df.iterrows():
    print(f"\n🌟 正在处理第 {idx+1} 个账户")

    try:
        # 取出Excel里这一行的数据
        cf = str(row["CODICE FISCALE"])
        pin = str(row["CODICE PIN"])
        old_pwd = str(row["PASSWORD INIZIALE"])
        new_pwd = str(row["NUOVA PASSWORD"])

        print(f"👉 Codice Fiscale: {cf}")

        # 打开税务局密码重置网页
        driver.get("https://telematici.agenziaentrate.gov.it/Abilitazione/RipristinaPassword/IRipristinaPassword.jsp")
        time.sleep(2)  # 等网页加载

        # 自动填写表单
        driver.find_element(By.NAME, "codFisc").send_keys(cf)
        driver.find_element(By.NAME, "pincode").send_keys(pin)
        driver.find_element(By.NAME, "oldpw").send_keys(old_pwd)
        driver.find_element(By.NAME, "newpw").send_keys(new_pwd)
        driver.find_element(By.NAME, "newpwf").send_keys(new_pwd)

        # ✅ 点击提交按钮
        print("👉 点击 INVIA")
        driver.find_element(By.CSS_SELECTOR, "input[type='submit']").click()

        print(f"✅ {cf} 已提交完成！")
        time.sleep(WAIT_BETWEEN_ACCOUNTS)

        page_html = driver.page_source

        if "Ripristino password effettuato con successo" in page_html:
            print(f"✅ {cf} 密码更新成功")
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            log_file.write(f"{timestamp:<20}  {cf:<25}\tSUCCESS\n")
        else:
            print(f"⚠️ {cf} 密码更新失败")
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            log_file.write(f"{timestamp}\t{cf}\tFAIL\n")

    except Exception as e:
        print(f"❌ 出错：{e}")
        time.sleep(2)

log_file.close()

# ✅ ④ 结束后关闭浏览器
driver.quit()
input("\n✅ 所有账号处理完毕，浏览器已关闭。\n\n请按 Enter 退出...")

