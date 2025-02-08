import requests
from bs4 import BeautifulSoup
import time
import csv
from urllib.parse import urljoin
import logging

# ロギングの設定
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

class MangaScraper:
    def __init__(self):
        self.base_url = "https://manba.co.jp/keyword_tags/{}/boards"
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }

    def check_page_exists(self, url):
        try:
            response = requests.get(url, headers=self.headers)
            # ステータスコードが200（成功）でかつ、コンテンツが存在する場合
            if response.status_code == 200 and 'boards-main-module' in response.text:
                return True, response.text
            return False, None
        except requests.RequestException as e:
            logging.error(f"Error accessing {url}: {str(e)}")
            return False, None

    def parse_manga_info(self, html):
        soup = BeautifulSoup(html, 'html.parser')
        manga_list = []
        
        for manga_div in soup.find_all('div', class_='boards-main'):
            try:
                # タイトル取得
                title_element = manga_div.find('h3', class_='board-title')
                if not title_element or not title_element.find('a'):
                    continue
                title = title_element.find('a').text.strip()
                
                # 著者取得
                authors = []
                author_elements = manga_div.find_all('div', class_='author-name')
                for author in author_elements:
                    if author.find('a'):
                        authors.append(author.find('a').text.strip())
                authors_str = ', '.join(authors)
                
                # 巻数取得
                books_count = manga_div.find('div', class_='books-count')
                volumes = books_count.find('a').text.strip() if books_count and books_count.find('a') else "不明"
                
                # 詳細URL取得
                detail_url = None
                if title_element.find('a').get('href'):
                    detail_url = urljoin('https://manba.co.jp', title_element.find('a').get('href'))
                
                manga_list.append({
                    'タイトル': title,
                    '著者': authors_str,
                    '巻数': volumes,
                    'URL': detail_url
                })
                
            except Exception as e:
                logging.error(f"Error parsing manga entry: {str(e)}")
                continue
        
        return manga_list

    def scrape_and_save(self, start_id=1, end_id=50, output_file='manga_data.csv'):
        manga_count = 0
        # BOM付きUTF-8で保存し、ヘッダーも日本語に変更
        with open(output_file, 'w', newline='', encoding='utf-8-sig') as csvfile:
            fieldnames = ['タイトル', '著者', '巻数', 'URL']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            
            for id in range(start_id, end_id + 1):
                url = self.base_url.format(id)
                logging.info(f"Checking ID {id}...")
                
                # ページの存在確認
                exists, html_content = self.check_page_exists(url)
                
                if exists:
                    logging.info(f"Found valid page for ID {id}, scraping...")
                    manga_list = self.parse_manga_info(html_content)
                    for manga in manga_list:
                        writer.writerow(manga)
                        manga_count += 1
                        logging.info(f"Saved: {manga['タイトル']}")
                else:
                    logging.info(f"No valid page found for ID {id}, skipping...")
                
                # 最低1秒の間隔を確保
                time.sleep(1)
        
        logging.info(f"Scraping completed! Total manga entries saved: {manga_count}")

def main():
    scraper = MangaScraper()
    scraper.scrape_and_save()

if __name__ == "__main__":
    main()