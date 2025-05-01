import json
import time
import os
import logging
import sys
import uuid, tempfile, shutil, atexit
import traceback
import subprocess
from selenium.webdriver.edge.options import Options as EdgeOptions
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import SessionNotCreatedException

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

def new_user_data_dir():
    return os.path.join(
        tempfile.gettempdir(),
        f"engage_profile_{uuid.uuid4().hex}"
    )

def build_options():
    opts = EdgeOptions()
    if HEADLESS:
        opts.add_argument("--headless=new")
        opts.add_argument("--window-size=1920,1080")
    else:
        opts.add_argument("--start-maximized")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-extensions")
    opts.add_argument("--remote-debugging-port=0")
    return opts

def create_driver():
    global user_data_dir
    # Edge の残骸プロセスを必ず殺してから試す
    subprocess.run(
        "taskkill /F /IM msedge.exe /T",
        shell=True,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL
    )

    user_data_dir = os.path.join(tempfile.gettempdir(),
                               f"engage_profile_{uuid.uuid4().hex}")
    opts = build_options()
    opts.add_argument(f"--user-data-dir={user_data_dir}")
    time.sleep(1)
    logger.info(f"user_data_dir(child) = {user_data_dir}")
    try:
        drv = webdriver.Edge(
            service=Service(EdgeChromiumDriverManager().install(),
                            log_path=os.devnull),
            options=opts
        )
        return drv
    except Exception:
        # 失敗したらフォルダを掃除して投げ直す
        shutil.rmtree(user_data_dir, ignore_errors=True)
        raise

for attempt in range(3):
    try:
        driver = create_driver()
        break
    except SessionNotCreatedException as e:
        logger.warning(f"プロファイル競合で再試行 {attempt+1}/3: {e}")
        subprocess.run(
            "taskkill /F /IM msedge.exe /T",
            shell=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )
        time.sleep(2)
else:
    logger.warning("フォールバック: user-data-dir なしで Edge を起動します")
    opts = build_options()          # この opts には user-data-dir を付けない
    driver = webdriver.Edge(
        service=Service(EdgeChromiumDriverManager().install(),
                        log_path=os.devnull),
        options=opts
    )

# driver が確定したあとに wait を作成
wait = WebDriverWait(driver, 10)

def login():
    logger.info("ログイン処理開始")
    driver.get("https://en-gage.net/company/manage/")
    if "login" in driver.current_url:
        wait.until(EC.presence_of_element_located((By.ID, "loginID"))).send_keys(EMAIL)
        driver.find_element(By.ID, "password").send_keys(PASSWORD)
        driver.find_element(By.ID, "login-button").click()
    logger.info("ログイン処理完了")

def close_popups(driver, wait):
    logger.info("ポップアップ閉じ処理開始")
    try:
        retry_outer = 0
        while retry_outer < 5:
            close_buttons = driver.find_elements(By.CLASS_NAME, "js_modalClose")
            if not close_buttons:
                break
            for btn in close_buttons:
                retry_inner = 0
                while retry_inner < 3:
                    try:
                        driver.execute_script("arguments[0].click();", btn)
                        time.sleep(1)
                        break
                    except Exception as e:
                        retry_inner += 1
                        logger.warning(f"✖クリックリトライ{retry_inner}回目失敗: {e}")
                        time.sleep(1)
                else:
                    logger.warning("✖ボタン押下3回失敗、次へ進みます。")
            retry_outer += 1
            time.sleep(1)
        else:
            logger.warning("ポップアップ閉じ処理を5回リトライしてもボタンが残っていました。")
        logger.info("ポップアップ閉じ処理完了")
    except Exception as e:
        logger.warning(f"ポップアップ閉じ処理中にエラー: {e}")

def select_job(driver, wait):
    logger.info("求人選択処理開始")
    job_select_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div.switch.js_modalOpen")))
    driver.execute_script("arguments[0].click();", job_select_button)
    time.sleep(1)

    job_elements = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a.js_optionWorkId")))
    job_titles = ["選択しない"] + [job.text for job in job_elements]

    if not job_titles:
        logger.error("求人リストが取得できませんでした。処理を終了します。")
        driver.quit()
        sys.exit(1)

    print("\n【選択可能な求人一覧】")
    for idx, title in enumerate(job_titles, start=0):  # 0スタート
        print(f"{idx}: {title}")

    while True:
        try:
            selected_idx = int(input("\n選択する求人の番号を入力してください: "))
            if 0 <= selected_idx < len(job_titles):
                break
            else:
                print("無効な番号です。もう一度入力してください。")
        except ValueError:
            print("数字を入力してください。")

    if selected_idx != 0:
        selected_element = job_elements[selected_idx - 1]  # インデックスずらす
        driver.execute_script("arguments[0].click();", selected_element)
        logger.info(f"求人選択完了: {job_titles[selected_idx]}")
    else:
        logger.info("求人選択スキップ（選択しない）")

    time.sleep(1)

def select_dropdowns():
    logger.info("ドロップダウン選択開始")
    pref_select = Select(driver.find_element(By.ID, "md_select-candidatePrefecture"))
    pref_options = []
    for option in pref_select.options:
        if option.text == "---":
            pref_options.append("選択しない")  # --- を 選択しない に置換
        else:
            pref_options.append(option.text)

    print("【現住所の選択肢】")
    for idx, text in enumerate(pref_options, start=0):  # 0スタート
        print(f"{idx}: {text}")

    while True:
        try:
            pref_idx = int(input("→ 番号を入力してください（例：0）："))
            if 0 <= pref_idx < len(pref_options):
                if pref_idx != 0:
                    pref_select.select_by_visible_text(pref_options[pref_idx])
                    logger.info(f"現住所選択：{pref_idx}:{pref_options[pref_idx]}")
                else:
                    logger.info("現住所選択スキップ（選択しない）")
                break
            else:
                print("無効な番号です。もう一度入力してください。")
        except ValueError:
            print("無効な入力です。数字で入力してください。")

    job_select = Select(driver.find_element(By.ID, "md_select-candidateOccupation"))
    job_options = []
    for option in job_select.options:
        if option.text == "---":
            job_options.append("選択しない")  # --- を 選択しない に置換
        else:
            job_options.append(option.text)

    print("【経験職種の選択肢】")
    for idx, text in enumerate(job_options, start=0):  # 0スタート
        print(f"{idx}: {text}")

    while True:
        try:
            job_idx = int(input("→ 番号を入力してください（例：0）："))
            if 0 <= job_idx < len(job_options):
                if job_idx != 0:
                    job_select.select_by_visible_text(job_options[job_idx])
                    logger.info(f"経験職種選択：{job_idx}:{job_options[job_idx]}")
                else:
                    logger.info("経験職種選択スキップ（選択しない）")
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
                print(f"対象候補者（{age}歳）処理開始")
                profile_buttons = driver.find_elements(By.XPATH, '//a[contains(@class, "md_btn--detail")]')
                if not profile_buttons:
                    logger.warning("プロフィールボタンが見つかりません。スキップ。")
                    continue
                driver.execute_script("arguments[0].scrollIntoView(true);", profile_buttons[0])
                driver.execute_script("arguments[0].click();", profile_buttons[0])
                logger.info("プロフィールを確認ボタン押下（JavaScriptクリック）")

                time.sleep(2)
                approach_buttons = driver.find_elements(By.XPATH, '//a[contains(@class, "js_candidateApproach")]')
                if approach_buttons:
                     driver.execute_script("arguments[0].click();", approach_buttons[0])
                     approach_count += 1
                     logger.info("会ってみたいボタン押下（JavaScript実行）")
                else:
                     logger.warning("会ってみたいボタンが見つかりませんでした。")
                time.sleep(1)

                body = driver.find_element(By.TAG_NAME, "body")
                ActionChains(driver).move_to_element_with_offset(body, 10, 10).click().perform()
                time.sleep(1)
                logger.info("プロフィールポップアップ閉じ完了")
        except Exception as e:
            logger.warning(f"候補者処理中エラー: {e}")

# 終了時に後片付け（1 回だけ登録）
def _cleanup():
    if 'user_data_dir' in globals() and os.path.exists(user_data_dir):
        shutil.rmtree(user_data_dir, ignore_errors=True)

# 実行開始
try:
    logger.info("自動処理開始")
    login()
    close_popups(driver, wait)
    select_job(driver, wait)
    select_dropdowns()
    process_candidates()
    logger.info(f"会ってみたいボタン押下人数: {approach_count}人")
    print(f"会ってみたいボタン押下人数: {approach_count}人")
    logger.info("自動処理完了")
    atexit.register(_cleanup)
    driver.quit()
except Exception as e:
    logger.error("FATAL ERROR: %s", e)
    logger.error(traceback.format_exc())
    # エラー内容をコンソールに残す
    input("スタックトレース出力")
    raise