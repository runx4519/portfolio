import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import threading
import time
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.common.action_chains import ActionChains
from datetime import timedelta

class EngageApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Engage 自動化ツール")
        self.geometry("600x700")
        self.job_titles = []
        self.load_config_options()
        self.create_variables()
        self.create_first_screen()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def load_config_options(self):
        with open("config_options.json", "r", encoding="utf-8") as f:
            data = json.load(f)
            self.pref_options = list(data["prefectures"].values())
            self.occ_options = list(data["occupations"].values())

    def create_variables(self):
        self.company_name = tk.StringVar()
        self.email = tk.StringVar()
        self.password = tk.StringVar()
        self.job1 = tk.StringVar()
        self.job2 = tk.StringVar()
        self.job3 = tk.StringVar()
        self.pref1 = tk.StringVar()
        self.pref2 = tk.StringVar()
        self.pref3 = tk.StringVar()
        self.occ1 = tk.StringVar()
        self.occ2 = tk.StringVar()
        self.occ3 = tk.StringVar()
        self.max_age = tk.StringVar()
        self.date_from = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        self.date_to = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))

    def create_first_screen(self):
        frame = ttk.Frame(self)
        frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="会社名:").pack(anchor=tk.W)
        ttk.Entry(frame, textvariable=self.company_name).pack(fill=tk.X)

        ttk.Label(frame, text="メールアドレス:").pack(anchor=tk.W)
        ttk.Entry(frame, textvariable=self.email).pack(fill=tk.X)

        ttk.Label(frame, text="パスワード:").pack(anchor=tk.W)
        ttk.Entry(frame, textvariable=self.password, show="*").pack(fill=tk.X)

        ttk.Button(frame, text="ログイン", command=self.initialize_driver).pack(pady=10)

    def initialize_driver(self):
        opts = EdgeOptions()
        opts.add_argument("--start-maximized")
        #opts.add_argument("--headless=new")
        self.driver = webdriver.Edge(service=Service(EdgeChromiumDriverManager().install()), options=opts)
        self.wait = WebDriverWait(self.driver, 5)
        self.driver.get("https://en-gage.net/company/manage/")

        if "login" in self.driver.current_url:
            self.wait.until(EC.presence_of_element_located((By.ID, "loginID"))).send_keys(self.email.get())
            self.driver.find_element(By.ID, "password").send_keys(self.password.get())
            self.driver.find_element(By.ID, "login-button").click()

        try:
            self.driver.get("https://en-gage.net/company/manage/")
            self.log_window("リロードでポップアップ抑制を試行中...")
            time.sleep(2)

            try:
                modal_close = self.driver.find_element(By.CSS_SELECTOR, "a.js_modalX.js_modalXChain")
                self.driver.execute_script("arguments[0].click();", modal_close)
                self.log_window("ポップアップを js_modalX で閉じました。")
                time.sleep(1)
            except Exception:
                self.log_window("js_modalX が見つかりませんでした（ポップアップ非表示とみなす）")

            self.close_popups()
            self.get_job_titles()
            self.create_second_screen()

        except Exception as e:
            messagebox.showerror("エラー", f"ログイン中にエラーが発生しました: {e}")

    def close_popups(self):
        self.log_window("ポップアップ閉じ処理開始")
        try:
            close_buttons = self.driver.find_elements(By.CLASS_NAME, "js_modalClose")
            if not close_buttons:
                self.log_window("ポップアップは検出されませんでした（スキップ）")
                return  # 無駄なループを回避

            retry_outer = 0
            while retry_outer < 5:
                for btn in close_buttons:
                    retry_inner = 0
                    while retry_inner < 3:
                        try:
                            self.driver.execute_script("arguments[0].scrollIntoView(true);", btn)
                            self.driver.execute_script("arguments[0].style.pointerEvents = 'auto';", btn)
                            self.driver.execute_script("arguments[0].click();", btn)
                            time.sleep(1)
                            break
                        except Exception as e:
                            retry_inner += 1
                            self.log_window(f"✖クリックリトライ{retry_inner}回目失敗: {e}")
                            time.sleep(1)
                    else:
                        self.log_window("✖ボタン押下3回失敗、次へ進みます。")
                retry_outer += 1
                time.sleep(1)
            else:
                self.log_window("ポップアップ閉じ処理を5回リトライしてもボタンが残っていました。")
            self.log_window("ポップアップ閉じ処理完了")
        except Exception as e:
            self.log_window(f"ポップアップ閉じ処理中にエラー: {e}")

            self.driver.execute_script("""
                const modals = document.querySelectorAll(
                    '.modal, .modalContainer, .js_modalXChain, .js_modalClose, .overlay, [class*="modal"]'
                );
                modals.forEach(el => {
                    // js_modalOpen（求人ボタン）だけは残す
                    if (!el.classList.contains('js_modalOpen')) {
                        el.style.display = 'none';
                        el.style.pointerEvents = 'none';
                        el.remove();
                    }
                });
            """)

            self.log_window("モーダル非表示成功")
        except Exception as e:
            self.log_window(f"モーダル非表示失敗: {e}")

    def get_job_titles(self):
        try:
            self.close_popups()

            # 背面レイヤーが残っていないか確認して消えるまで待つ
            try:
                WebDriverWait(self.driver, 5).until_not(
                    EC.presence_of_element_located((By.CLASS_NAME, "js_modalXChain"))
                )
            except:
                pass

            # 求人モーダルボタンを確実にクリック（scrollIntoView付き）
            modal_button = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "div.switch.js_modalOpen"))
            )
            self.driver.execute_script("arguments[0].scrollIntoView(true);", modal_button)
            self.driver.execute_script("arguments[0].click();", modal_button)
            time.sleep(1)

            # 求人タイトル一覧を取得
            job_elements = self.wait.until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a.js_optionWorkId"))
            )
            self.job_titles = ["選択しない"] + [job.text for job in job_elements]

            # 求人情報取得処理のあと
            print("求人タイトル取得完了。モーダル閉じ開始")
            # 確認後すぐ処理継続
            time.sleep(1)

            # 求人情報ポップアップは F5 リロードで閉じる（確実に消えるならこの方法が安定）
            try:
                self.driver.get("https://en-gage.net/company/manage/")
                self.log_window("求人情報ポップアップをリロードで強制非表示にしました。")
                time.sleep(2)
            except Exception as e:
                self.log_window(f"リロードによる求人情報ポップアップ非表示失敗: {e}")
        except Exception as e:
            messagebox.showerror("エラー", f"求人情報の取得に失敗しました: {e}")
        # モーダルを安定して閉じる
        self.close_popups()
        # 2画面目を表示（求人情報取得完了後）
        self.create_second_screen()

    def create_second_screen(self):
        for widget in self.winfo_children():
            widget.destroy()

        frame = ttk.Frame(self)
        frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        for i, var in enumerate([self.job1, self.job2, self.job3], start=1):
            ttk.Label(frame, text=f"求人情報 {i}:").pack(anchor=tk.W)
            ttk.Combobox(frame, textvariable=var, values=self.job_titles).pack(fill=tk.X)

        for i, var in enumerate([self.pref1, self.pref2, self.pref3], start=1):
            ttk.Label(frame, text=f"現住所 {i}:").pack(anchor=tk.W)
            ttk.Combobox(frame, textvariable=var, values=self.pref_options).pack(fill=tk.X)

        for i, var in enumerate([self.occ1, self.occ2, self.occ3], start=1):
            ttk.Label(frame, text=f"経験職種 {i}:").pack(anchor=tk.W)
            ttk.Combobox(frame, textvariable=var, values=self.occ_options).pack(fill=tk.X)

        ttk.Label(frame, text="候補者年齢（以下）:").pack(anchor=tk.W)
        ttk.Entry(frame, textvariable=self.max_age).pack(fill=tk.X)

        ttk.Label(frame, text="実行期間 From（例: 2025-05-08）:").pack(anchor=tk.W)
        ttk.Entry(frame, textvariable=self.date_from).pack(fill=tk.X)

        ttk.Label(frame, text="実行期間 To（例: 2025-05-08）:").pack(anchor=tk.W)
        ttk.Entry(frame, textvariable=self.date_to).pack(fill=tk.X)

        ttk.Button(frame, text="実行", command=self.start_automation_thread).pack(pady=10)
        self.log_box = tk.Text(frame, height=15)
        self.log_box.pack(fill=tk.BOTH, expand=True)

    def log_window(self, message):
        now = datetime.now().strftime("%H:%M:%S")
        if hasattr(self, "log_box"):
            self.log_box.insert(tk.END, f"[{now}] {message}\n")
            self.log_box.see(tk.END)
        else:
            print(f"[{now}] {message}")  # 初期画面では print にフォールバック

    def validate_inputs(self):
        if not self.company_name.get().strip():
            messagebox.showwarning("入力エラー", "会社名は必須です")
            return False
        if not self.email.get().strip():
            messagebox.showwarning("入力エラー", "メールアドレスは必須です")
            return False
        if not self.password.get().strip():
            messagebox.showwarning("入力エラー", "パスワードは必須です")
            return False

        if not self.date_from.get().strip():
            messagebox.showwarning("入力エラー", "実行期間Fromは必須です")
            return False
        if not self.date_to.get().strip():
            messagebox.showwarning("入力エラー", "実行期間Toは必須です")
            return False

        try:
            age = int(self.max_age.get().strip())
            if age <= 0:
                raise ValueError
        except:
            messagebox.showwarning("入力エラー", "候補者年齢は正の整数で入力してください")
            return False

        if all(j in [""] for j in [self.job1.get(), self.job2.get(), self.job3.get()]):
            messagebox.showwarning("入力エラー", "求人情報はいずれか1つ以上選択してください")
            return False
        if all(p in [""] for p in [self.pref1.get(), self.pref2.get(), self.pref3.get()]):
            messagebox.showwarning("入力エラー", "現住所はいずれか1つ以上選択してください")
            return False
        if all(o in [""] for o in [self.occ1.get(), self.occ2.get(), self.occ3.get()]):
            messagebox.showwarning("入力エラー", "経験職種はいずれか1つ以上選択してください")
            return False

        return True

    def start_automation_thread(self):
        if not self.validate_inputs():
            return
        self.disable_inputs()
        thread = threading.Thread(target=self.run_automation_wrapper, daemon=True)
        thread.start()

    def run_automation(self):
        self.log_window("処理を開始しました。")
        try:    
            from_time = datetime.strptime(self.date_from.get().strip() + " 00:00:00", "%Y-%m-%d %H:%M:%S")
            to_time = datetime.strptime(self.date_to.get().strip() + " 23:59:59", "%Y-%m-%d %H:%M:%S")
            now = datetime.now()
            if now < from_time:
                wait_minutes = (from_time - now).total_seconds() / 60
                self.log_window(f"開始時刻まで {int(wait_minutes)} 分待機します...")
                time.sleep(wait_minutes * 60)

            jobs = [self.job1.get(), self.job2.get(), self.job3.get()]
            prefs = [self.pref1.get(), self.pref2.get(), self.pref3.get()]
            occs  = [self.occ1.get(), self.occ2.get(), self.occ3.get()]

            loop_count = 0
            while datetime.now() <= to_time:
                self.log_window(f"{loop_count + 1}回目のループ処理を開始します")
                for i in range(3):
                    job = jobs[i]
                    pref = prefs[i]
                    occ = occs[i]
                    if "選択しない" in (job, pref, occ):
                        self.log_window(f"{i+1}回目: 条件が不完全なためスキップ（{job}, {pref}, {occ}）")
                        continue

                    self.log_window(f"{i+1}回目: 求人={job}, 現住所={pref}, 職種={occ} で処理開始")
                    self.run_condition_set(job, pref, occ)
                    self.log_window(f"{i+1}回目: 処理完了")

                loop_count += 1
                if datetime.now() + timedelta(minutes=20) > to_time:
                    break
                self.log_window("20分待機します...")
                time.sleep(20 * 60)
            self.log_window("実行期間終了。処理を終了します。")
        except Exception as e:
            self.log_window(f"処理中にエラー: {e}")
        self.log_window("処理を終了しました。")

    def run_condition_set(self, job_name, pref_name, occ_name):
        self.log_window(f"▶ 条件: {job_name} × {pref_name} × {occ_name} を処理中...")
        try:
            # 求人選択
            self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div.switch.js_modalOpen"))).click()
            time.sleep(1)
            job_elements = self.wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a.js_optionWorkId")))
            for job in job_elements:
                if job.text == job_name:
                    self.driver.execute_script("arguments[0].click();", job)
                    break
            time.sleep(1)

            # 現住所と経験職種の選択
            pref_select = Select(self.driver.find_element(By.ID, "md_select-candidatePrefecture"))
            pref_select.select_by_visible_text(pref_name)
            time.sleep(0.5)

            occ_select = Select(self.driver.find_element(By.ID, "md_select-candidateOccupation"))
            occ_select.select_by_visible_text(occ_name)
            time.sleep(0.5)

            # 絞り込みボタン押下
            refine_btn = self.driver.find_element(By.ID, "js_candidateRefinement")
            self.driver.execute_script("arguments[0].click();", refine_btn)
            time.sleep(2)

            # さらに読み込む最大まで
            for _ in range(9):
                try:
                    show_more = WebDriverWait(self.driver, 3).until(
                        EC.element_to_be_clickable((By.ID, "js_candidateShowMore"))
                    )
                    self.driver.execute_script("arguments[0].click();", show_more)
                    time.sleep(2)
                except:
                    break

            # 候補者一覧から年齢条件に合う人を抽出して会ってみたい
            candidates = self.driver.find_elements(By.XPATH, '//div[@class="main"]')
            for candidate in candidates:
                self.log_window("候補者プロフィールを確認中...")
                try:
                    age_elements = candidate.find_elements(By.XPATH, './/span[contains(text(), "歳")]')
                    if not age_elements:
                        continue
                    age_text = age_elements[0].text
                    age = int(age_text.replace("歳", ""))
                    if age > int(self.max_age.get()):
                        continue

                    profile_button = candidate.find_element(By.XPATH, './/a[contains(@class, "md_btn--detail")]')
                    self.driver.execute_script("arguments[0].click();", profile_button)
                    time.sleep(2)

                    approach_btns = self.driver.find_elements(By.XPATH, '//a[contains(@class, "js_candidateApproach")]')
                    if approach_btns:
                    #    self.driver.execute_script("arguments[0].click();", approach_btns[0])
                        self.log_window("会ってみたいボタン押下")
                    else:
                        self.log_window("会ってみたいボタンなし")

                    # ポップアップ閉じる
                    body = self.driver.find_element(By.TAG_NAME, "body")
                    ActionChains(self.driver).move_to_element_with_offset(body, 10, 10).click().perform()
                    time.sleep(1)
                except Exception as e:
                    self.log_window(f"候補者処理エラー: {e}")
        except Exception as e:
            self.log_window(f"run_condition_set エラー: {e}")

    def disable_inputs(self):
        for child in self.winfo_children():
            for widget in child.winfo_children():
                try:
                    widget.configure(state="disabled")
                except:
                    pass

    def enable_inputs(self):
        for child in self.winfo_children():
            for widget in child.winfo_children():
                try:
                    widget.configure(state="normal")
                except:
                    pass

    def run_automation_wrapper(self):
        try:
            self.run_automation()
        finally:
            self.enable_inputs()

    def on_closing(self):
        try:
            if hasattr(self, "driver"):
                self.driver.quit()  # Edgeが生きていれば閉じる
        except:
            pass
        self.destroy()

if __name__ == "__main__":
    app = EngageApp()
    app.mainloop()