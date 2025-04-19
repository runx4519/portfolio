from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import csv
from datetime import datetime

def get_google_suggestions(driver, keyword):
    driver.get('https://www.google.co.jp')
    search_box = driver.find_element(By.NAME, 'q')
    search_box.clear()
    search_box.send_keys(keyword)
    time.sleep(1.5)
    suggestion_elements = driver.find_elements(By.CSS_SELECTOR, 'ul[role="listbox"] li span')
    suggestions = [el.text for el in suggestion_elements if el.text]
    return suggestions[:9]

def main():
    start_time = datetime.now()
    print(f"[開始] {start_time.strftime('%Y-%m-%d %H:%M:%S')}")

    input_file = 'keyword.txt'
    timestamp = start_time.strftime('%Y%m%d%H%M%S')
    output_file = f'suggestions_{timestamp}.csv'

    with open(input_file, 'r', encoding='utf-8') as f:
        keywords = [line.strip() for line in f if line.strip()]

    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    driver = webdriver.Chrome(options=options)

    with open(output_file, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow(['検索キーワード', 'サジェスト'])

        for keyword in keywords:
            try:
                print(f"検索中: {keyword}")
                suggestions = get_google_suggestions(driver, keyword)
                if suggestions:
                    writer.writerow([keyword, suggestions[0]])
                    for s in suggestions[1:]:
                        writer.writerow(['', s])
                else:
                    writer.writerow([keyword, ''])
            except Exception as e:
                print(f"エラー発生: {keyword} → {e}")
                writer.writerow([keyword, '取得失敗'])

    driver.quit()

    end_time = datetime.now()
    print(f"[終了] {end_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"\nCSV出力完了：{output_file}")

if __name__ == '__main__':
    main()
