from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time

def setup_driver():
    """Seleniumのドライバーをセットアップする"""
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--window-size=1920,1080')
    
    service = ChromeService(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver

def get_exhibitor_info(exhibitor):
    """各出展者の情報を取得する"""
    try:
        # 会社名を取得
        company_name = exhibitor.find_element(By.CSS_SELECTOR, "h3.exhibitor-name").text.strip()
        
        # Webリンクを取得
        web_link = exhibitor.find_element(By.CSS_SELECTOR, "h3.exhibitor-name").find_element(By.XPATH, "..").get_attribute("href")
        
        # ブランド名を取得
        try:
            brands = exhibitor.find_element(By.CSS_SELECTOR, "div.brands p").text.strip()
        except NoSuchElementException:
            brands = ""
            
        # 登録カテゴリーを取得
        try:
            categories = exhibitor.find_element(By.CSS_SELECTOR, "span.pps-tags").text.strip()
        except NoSuchElementException:
            categories = ""
            
        return {
            "会社名": company_name,
            "web_link": web_link,
            "ブランド名": brands,
            "登録カテゴリー": categories
        }
    except StaleElementReferenceException:
        print("要素が古くなったため、スキップします")
        return None
    except Exception as e:
        print(f"Error extracting exhibitor info: {e}")
        return None

def scroll_and_get_exhibitors(driver, target_count=1987):
    """スクロールしながら出展者情報を取得する"""
    exhibitors_data = []
    last_height = driver.execute_script("return document.body.scrollHeight")
    scroll_attempts = 0
    max_scroll_attempts = 50  # 最大スクロール試行回数
    
    while len(exhibitors_data) < target_count and scroll_attempts < max_scroll_attempts:
        try:
            # 現在表示されている出展者要素を取得
            exhibitor_elements = driver.find_elements(
                By.CSS_SELECTOR, 
                "#exhibitor-directory > div > div:nth-child(1) > div > div.filter-results > div > div > ul > div"
            )
            
            # 新しい出展者情報を取得
            current_count = len(exhibitors_data)
            for exhibitor in exhibitor_elements[current_count:]:
                try:
                    # 要素までスクロール
                    driver.execute_script("arguments[0].scrollIntoView(true);", exhibitor)
                    time.sleep(0.5)
                    
                    info = get_exhibitor_info(exhibitor)
                    if info and info not in exhibitors_data:
                        exhibitors_data.append(info)
                        if len(exhibitors_data) % 50 == 0:
                            print(f"取得済み出展者数: {len(exhibitors_data)}")
                            # 途中経過を保存
                            pd.DataFrame(exhibitors_data).to_csv(
                                f'manufacturing_world_exhibitors_temp_{len(exhibitors_data)}.csv',
                                index=False, encoding='utf-8-sig'
                            )
                except StaleElementReferenceException:
                    continue
            
            # ページ下部まで少しずつスクロール
            for _ in range(3):
                driver.execute_script(
                    "window.scrollTo(0, window.scrollY + window.innerHeight/2);"
                )
                time.sleep(1)
            
            # 新しい高さを取得
            new_height = driver.execute_script("return document.body.scrollHeight")
            
            # スクロールしても高さが変わらない場合
            if new_height == last_height:
                scroll_attempts += 1
                print(f"スクロール試行回数: {scroll_attempts}")
                
                # 少し待って再度スクロールを試みる
                time.sleep(2)
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(2)
            else:
                scroll_attempts = 0  # 高さが変わったらカウントをリセット
                
            last_height = new_height
            
        except Exception as e:
            print(f"スクロール中にエラーが発生: {e}")
            scroll_attempts += 1
            time.sleep(2)
    
    return exhibitors_data

def scrape_manufacturing_world():
    """展示会サイトをスクレイピングするメイン関数"""
    url = "https://www.manufacturing-world.jp/tokyo/ja-jp/search/2024/directory.html?#/"
    driver = setup_driver()
    
    try:
        driver.get(url)
        print("ページの読み込みを待機中...")
        time.sleep(5)
        
        # ページ上部に表示される可能性のあるモーダルやポップアップを閉じる処理を追加
        try:
            close_buttons = driver.find_elements(By.CSS_SELECTOR, "button[aria-label='Close']")
            for button in close_buttons:
                button.click()
                time.sleep(1)
        except:
            pass
        
        print("スクレイピングを開始します...")
        exhibitors_data = scroll_and_get_exhibitors(driver)
        
        # 最終データを保存
        df = pd.DataFrame(exhibitors_data)
        df.to_csv('manufacturing_world_exhibitors_final.csv', index=False, encoding='utf-8-sig')
        print(f"\nスクレイピング完了! 合計{len(exhibitors_data)}社の情報を取得しました。")
        
    except Exception as e:
        print(f"スクレイピング中にエラーが発生: {e}")
        # エラー発生時も途中までのデータを保存
        if 'exhibitors_data' in locals() and exhibitors_data:
            pd.DataFrame(exhibitors_data).to_csv(
                'manufacturing_world_exhibitors_error_backup.csv',
                index=False, encoding='utf-8-sig'
            )
    
    finally:
        driver.quit()

if __name__ == "__main__":
    scrape_manufacturing_world()
