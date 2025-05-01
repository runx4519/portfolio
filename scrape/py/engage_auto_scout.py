import json
import time
import os
import logging
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.common.action_chains import ActionChains

# ログ設定
log_dir = "logs"
os.makedirs(log_dir, exist_ok=True)
log_filename = os.path.join(log_dir, datetime.now().strftime("engage_log_%Y%m%d.log"))
logging.basicConfig(
    filename=log_filename,
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

# 設定ファイル読み込み
with open("config.json", "r", encoding="utf-8") as f:
    config = json.load(f)

EMAIL = config["email"]
PASSWORD = config["password"]
MAX_AGE = config.get("max_age", 35)
HEADLESS = config.get("headless", False)

# 会ってみたいボタン押下カウンター
approach_count = 0

# Edge起動設定
options = Options()
if HEADLESS:
    options.add_argument("--headless=new")
    options.add_argument("--window-size=1920,1080")
else:
    options.add_argument("--start-maximized")

driver = webdriver.Edge(service=Service(EdgeChromiumDriverManager().install()), options=options)
wait = WebDriverWait(driver, 10)

def login():
    logger.info("ログイン処理開始")
    driver.get("https://en-gage.net/company/manage/")
    if "login" in driver.current_url:
        wait.until(EC.presence_of_element_located((By.ID, "loginID"))).send_keys(EMAIL)
        driver.find_element(By.ID, "password").send_keys(PASSWORD)
        driver.find_element(By.ID, "login-button").click()
    logger.info("ログイン処理完了")

def close_popups():
    logger.info("ポップアップ閉じ処理開始")
    for i in range(5):
        close_buttons = driver.find_elements(By.CLASS_NAME, "js_modalX")
        visible_close_buttons = [btn for btn in close_buttons if btn.is_displayed()]
        if not visible_close_buttons:
            break
        for btn in visible_close_buttons:
            try:
                logger.info("xボタンをクリックしてポップアップ閉じる")
                btn.click()
                time.sleep(1)
            except Exception as e:
                logger.warning(f"xクリック失敗: {e}")
                body = driver.find_element(By.TAG_NAME, "body")
                ActionChains(driver).move_to_element_with_offset(body, 10, 10).click().perform()
                time.sleep(1)
    logger.info("ポップアップ閉じ処理完了")

def select_dropdowns():
    logger.info("ドロップダウン選択開始")
    pref_select = Select(driver.find_element(By.ID, "md_select-candidatePrefecture"))
    pref_options = [option.text for option in pref_select.options if option.text != "---"]
    print("【現住所の選択肢】")
    for idx, text in enumerate(pref_options, start=1):
        print(f"{idx}: {text}")
    while True:
        try:
            pref_idx = int(input("→ 番号を入力してください（例：1）："))
            if 1 <= pref_idx <= len(pref_options):
                pref_select.select_by_visible_text(pref_options[pref_idx - 1])
                logger.info(f"現住所選択：{pref_idx}:{pref_options[pref_idx - 1]}")
                break
            else:
                print("無効な番号です。もう一度入力してください。")
        except ValueError:
            print("無効な入力です。数字で入力してください。")

    job_select = Select(driver.find_element(By.ID, "md_select-candidateOccupation"))
    job_options = [option.text for option in job_select.options if option.text != "---"]
    print("【経験職種の選択肢】")
    for idx, text in enumerate(job_options, start=1):
        print(f"{idx}: {text}")
    while True:
        try:
            job_idx = int(input("→ 番号を入力してください（例：1）："))
            if 1 <= job_idx <= len(job_options):
                job_select.select_by_visible_text(job_options[job_idx - 1])
                logger.info(f"経験職種選択：{job_idx}:{job_options[job_idx - 1]}")
                break
            else:
                print("無効な番号です。もう一度入力してください。")
        except ValueError:
            print("無効な入力です。数字で入力してください。")

    element = wait.until(EC.presence_of_element_located((By.ID, "js_candidateRefinement")))
    driver.execute_script("arguments[0].click();", element)
    logger.info("絞り込みボタン押下（JavaScript実行）")

def process_candidates():
    global approach_count
    logger.info("候補者処理開始")
    time.sleep(3)
    candidates = driver.find_elements(By.XPATH, '//div[@class="main"]')
    if not candidates:
        logger.info("対象候補者が見つかりませんでした。処理終了。")
        return

    for candidate in candidates:
        try:
            age_elements = candidate.find_elements(By.XPATH, './/span[contains(text(), "歳")]')
            if not age_elements:
                logger.warning("年齢情報が見つかりません（年齢不明）。スキップ。")
                continue
            age_text = age_elements[0].text
            age = int(age_text.replace("歳", ""))
            if age <= MAX_AGE:
                logger.info(f"対象候補者（{age}歳）処理開始")
                profile_buttons = driver.find_elements(By.XPATH, '//a[contains(@class, "md_btn--detail")]')
                if not profile_buttons:
                    logger.warning("プロフィールボタンが見つかりません。スキップ。")
                    continue
                driver.execute_script("arguments[0].scrollIntoView(true);", profile_buttons[0])
                driver.execute_script("arguments[0].click();", profile_buttons[0])
                logger.info("プロフィールを確認ボタン押下（JavaScriptクリック）")

                time.sleep(2)
                # approach_buttons = driver.find_elements(By.XPATH, '//a[contains(@class, "js_candidateApproach")]')
                # if approach_buttons:
                #     driver.execute_script("arguments[0].click();", approach_buttons[0])
                approach_count += 1
                logger.info("会ってみたいボタン押下（JavaScript実行）")
                # else:
                #     logger.warning("会ってみたいボタンが見つかりませんでした。")
                time.sleep(1)

                body = driver.find_element(By.TAG_NAME, "body")
                ActionChains(driver).move_to_element_with_offset(body, 10, 10).click().perform()
                time.sleep(1)
                logger.info("プロフィールポップアップ閉じ完了")
        except Exception as e:
            logger.warning(f"候補者処理中エラー: {e}")

# 実行開始
logger.info("自動処理開始")
login()
close_popups()
select_dropdowns()
process_candidates()
logger.info(f"会ってみたいボタン押下人数: {approach_count}人")
print(f"会ってみたいボタン押下人数: {approach_count}人")
logger.info("自動処理完了")
driver.quit()