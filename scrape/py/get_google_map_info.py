import argparse
import sys
import datetime
import time
import csv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

def setup_driver(headless=False):
    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument("--headless=new")
    else:
        options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    return driver

def scroll_and_collect_links(driver):
    """
    結果一覧のコンテナ内をスクロールしながら、
    読み込まれる全リンクを重複なく取得する。
    リンク数に変化がなくなったらループ終了。
    """
    urls = set()
    time.sleep(5)
    result_container = driver.find_element(By.XPATH, "//div[contains(@aria-label, '結果')]")
    
    last_count = 0
    no_change_count = 0
    max_no_change = 3  # 連続してリンク数に変化がなければ終了
    
    for _ in range(50):
        driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", result_container)
        time.sleep(2)
        links = result_container.find_elements(By.XPATH, ".//a[contains(@class, 'hfpxzc')]")
        for link in links:
            href = link.get_attribute("href")
            if href:
                urls.add(href)
        if len(urls) == last_count:
            no_change_count += 1
            if no_change_count >= max_no_change:
                break
        else:
            no_change_count = 0
            last_count = len(urls)
    
    return list(urls)

def scrape_details(driver, urls):
    """
    各リンク先のページから、事務所名、ホームページ、住所を取得する。
    """
    results = []
    for url in urls:
        driver.get(url)
        time.sleep(5)
        try:
            name = driver.find_element(
                By.XPATH,
                "//div[contains(@class, 'W1neJ')]/span[contains(@class, 'iD2gKb W1neJ')]"
            ).text.strip()
        except:
            try:
                name = driver.find_element(
                    By.XPATH,
                    "//h1[contains(@class, 'DUwDvf lfPIob')]"
                ).text.strip()
            except:
                name = "なし"
        try:
            website = driver.find_element(By.XPATH, "//a[@data-item-id='authority']").get_attribute("href")
        except:
            website = "なし"
        try:
            address = driver.find_element(
                By.XPATH,
                "//button[@data-item-id='address']/div/div[2]/div"
            ).text.strip()
        except:
            address = "なし"
        
        results.append([name, website, address])
    return results

def save_to_csv(data, filename_prefix="get_result.info"):
    # 現在の日時をYYYYMMDDHHMMSS形式で取得（例: 20250313154500）
    timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    filename = f"{filename_prefix}_{timestamp}.csv"
    
    with open(filename, mode='w', encoding='utf-8', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(["事務所名", "ホームページ", "住所"])
        writer.writerows(data)
    
    print(f"CSVファイル '{filename}' に保存しました。")

def main():
    parser = argparse.ArgumentParser(description="Googleマップスクレイピング")
    parser.add_argument("--headless", action="store_true", help="ヘッドレスモードで実行する場合に指定")
    args = parser.parse_args()
    
    # 検索キーワードの入力（空の場合は処理終了）
    keyword = input("検索キーワードを入力してください: ").strip()
    if not keyword:
        print("検索キーワードが入力されなかったため、処理を終了します。")
        sys.exit(0)
    
    # 入力キーワードをURL用に変換（スペースを"+"に）
    search_url = "https://www.google.co.jp/maps/search/" + keyword.replace(" ", "+")
    
    # 処理開始時刻を表示
    start_time = datetime.datetime.now()
    print("処理開始時間:", start_time.strftime("%Y-%m-%d %H:%M:%S"))
    
    driver = setup_driver(headless=args.headless)
    print("スクレイピング開始...")
    driver.get(search_url)
    
    urls = scroll_and_collect_links(driver)
    print(f"{len(urls)} 件のリンクを取得")
    
    if urls:
        data = scrape_details(driver, urls)
        save_to_csv(data)
        print("データ保存完了")
    else:
        print("リンクが取得できませんでした。")
    
    driver.quit()
    
    # 処理終了時刻を表示
    end_time = datetime.datetime.now()
    print("処理終了時間:", end_time.strftime("%Y-%m-%d %H:%M:%S"))
    elapsed = end_time - start_time
    print("処理時間:", elapsed)

if __name__ == "__main__":
    main()
