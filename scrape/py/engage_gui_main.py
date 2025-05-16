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
from selenium.common.exceptions import TimeoutException

class EngageApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Engage 自動化ツール")
        self.job_titles = []

        self.job_combo_list = []
        self.pref_combo_list = []
        self.occ_combo_list = []

        self.load_config_options()
        self.create_variables()
        self.create_first_screen()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def load_config_options(self):
        with open("config_options.json", "r", encoding="utf-8") as f:
            data = json.load(f)

        # 名称リスト（プルダウン表示用）
        self.pref_names = list(data["prefectures"].values())
        self.occ_names = list(data["occupations"].values())

        # 名称 → ID 辞書（検索条件用）
        self.pref_name_id_map = {v: k for k, v in data["prefectures"].items()}
        self.occ_name_id_map = {v: k for k, v in data["occupations"].items()}

    def create_variables(self):
        self.company_name = tk.StringVar()
        self.email = tk.StringVar()
        self.password = tk.StringVar()
        self.job1 = tk.StringVar()
        self.job2 = tk.StringVar()
        self.job3 = tk.StringVar()
        self.job_vars = [self.job1, self.job2, self.job3]
        self.pref1 = tk.StringVar()
        self.pref2 = tk.StringVar()
        self.pref3 = tk.StringVar()
        self.pref_vars = [self.pref1, self.pref2, self.pref3]
        self.occ1 = tk.StringVar()
        self.occ2 = tk.StringVar()
        self.occ3 = tk.StringVar()
        self.occ_vars = [self.occ1, self.occ2, self.occ3]
        self.max_age = tk.StringVar()
        self.date_from = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        self.date_to = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))

    def create_first_screen(self):
        self.geometry("300x200")
        frame = ttk.Frame(self)
        frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        # 会社名
        row1 = ttk.Frame(frame)
        row1.pack(anchor=tk.W, pady=5)
        ttk.Label(row1, text="会社名:", width=10).pack(side=tk.LEFT)
        ttk.Entry(row1, textvariable=self.company_name, width=30).pack(side=tk.LEFT, padx=5)

        # メールアドレス
        row2 = ttk.Frame(frame)
        row2.pack(anchor=tk.W, pady=5)
        ttk.Label(row2, text="メールアドレス:", width=10).pack(side=tk.LEFT)
        ttk.Entry(row2, textvariable=self.email, width=30).pack(side=tk.LEFT, padx=5)

        # パスワード
        row3 = ttk.Frame(frame)
        row3.pack(anchor=tk.W, pady=5)
        ttk.Label(row3, text="パスワード:", width=10).pack(side=tk.LEFT)
        ttk.Entry(row3, textvariable=self.password, width=30, show="*").pack(side=tk.LEFT, padx=5)

        # ログインボタン
        ttk.Button(frame, text="ログイン", command=self.on_login_button_click).pack(pady=10)

    def on_login_button_click(self):
        if not self.validate_first_screen():
            return
        if not self.initialize_driver():
            return

        self.get_job_titles()
        self.create_second_screen()

    def initialize_driver(self):
        opts = EdgeOptions()
        opts.add_argument("--start-maximized")
        opts.add_argument("--headless=new")
        self.driver = webdriver.Edge(service=Service(EdgeChromiumDriverManager().install()), options=opts)
        self.wait = WebDriverWait(self.driver, 5)
        self.driver.get("https://en-gage.net/company/manage/")

        if "login" in self.driver.current_url:
            self.wait.until(EC.presence_of_element_located((By.ID, "loginID"))).send_keys(self.email.get())
            self.driver.find_element(By.ID, "password").send_keys(self.password.get())
            self.driver.find_element(By.ID, "login-button").click()

            # ログイン失敗チェック
            try:
                WebDriverWait(self.driver, 3).until(
                    EC.presence_of_element_located((By.ID, "login-error-area"))
                )
                messagebox.showerror("ログインエラー", "「メールアドレス」「パスワード」を正しく入力してください。")
                self.driver.quit()
                return False
            except TimeoutException:
                self.driver.get("https://en-gage.net/company/manage/")
                self.log_window("リロードでポップアップ抑制を試行中...")
                time.sleep(2)

                # 再リロードで両ポップアップをまとめて閉じる
                self.driver.get("https://en-gage.net/company/manage/")
                self.log_window("ポップアップ類をまとめてリロードで閉じました")
                time.sleep(2)

                return True

    def close_popups(self):
        self.log_window("ポップアップ非表示処理を実行中")
        try:
            self.driver.execute_script("""
                const modals = document.querySelectorAll(
                    '.modal, .modalContainer, .js_modalXChain, .js_modalClose, .overlay, [class*="modal"]'
                );
                modals.forEach(el => {
                    if (!el.classList.contains('js_modalOpen')) {
                        el.style.display = 'none';
                        el.style.pointerEvents = 'none';
                    }
                });
                document.body.style.overflow = 'auto';
            """)
            self.log_window("全ポップアップを非表示にしました")
        except Exception as e:
            self.log_window(f"ポップアップ非表示中にエラー: {e}")

    def get_job_titles(self):
        try:
            modal_button = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "div.switch.js_modalOpen"))
            )
            self.driver.execute_script("arguments[0].scrollIntoView(true);", modal_button)
            self.driver.execute_script("arguments[0].click();", modal_button)
            self.log_window("求人情報モーダルを開きました")
            time.sleep(1)

            job_elements = self.wait.until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a.js_optionWorkId"))
            )

            job_map = {}
            for job in job_elements:
                job_text = job.text.strip()
                job_id = job.get_attribute("data-work_id")
                if job_text and job_id:
                    title_line = job_text.splitlines()[0]
                    display_title = f"【{job_id}】{title_line}"
                    job_map[display_title] = job_id

            self.job_title_id_map = job_map
            self.job_titles = ["選択しない"] + list(job_map.keys())

            self.log_window(f"求人情報 {len(job_map)} 件取得完了")
            self.close_popups()

        except TimeoutException:
            self.log_window("求人情報取得に失敗しました（Timeout）")
            self.close_popups()

    def create_second_screen(self):
        self.geometry("1000x600")
        self.clear_screen()
        frame = ttk.Frame(self)
        frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        ttk.Label(frame, text=f"会社名：{self.company_name.get()}", font=("Arial", 14, "bold")).pack(pady=(10, 5))

        # 入力エリアの親Frame（grid配置）
        input_frame = ttk.Frame(frame)
        input_frame.pack(anchor="w")

        # 求人情報
        self.add_label_combobox_row(frame, "求人情報", self.job_vars, self.job_titles, self.job_combo_list)
        # 現住所
        self.add_label_combobox_row(frame, "現住所", self.pref_vars, self.pref_names, self.pref_combo_list)
        # 経験職種
        self.add_label_combobox_row(frame, "経験職種", self.occ_vars, self.occ_names, self.occ_combo_list)

        # 候補者年齢
        age_frame = ttk.Frame(frame)
        age_frame.pack(anchor="w", pady=5)
        ttk.Label(age_frame, text="候補者年齢（以下）：").pack(side=tk.LEFT)
        ttk.Entry(age_frame, textvariable=self.max_age, width=10).pack(side=tk.LEFT, padx=5)

        # 実行期間 From/To
        for label, var in [("実行期間 From（例: 2025-05-08）：", self.date_from), ("実行期間 To（例: 2025-05-08）：", self.date_to)]:
            f = ttk.Frame(frame)
            f.pack(anchor="w", pady=5)
            ttk.Label(f, text=label).pack(side=tk.LEFT)
            ttk.Entry(f, textvariable=var, width=15).pack(side=tk.LEFT, padx=5)

        # 実行ボタン
        ttk.Button(frame, text="実行", command=self.start_automation_thread).pack(pady=10)

        # ログ出力
        self.log_box = tk.Text(frame, height=15)
        self.log_box.pack(fill=tk.BOTH, expand=True)

    def log_window(self, message):
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        if hasattr(self, "log_box"):
            self.log_box.config(state="normal")
            self.log_box.insert(tk.END, f"[{now}] {message}\n")
            self.log_box.see(tk.END)
            self.log_box.config(state="disabled") 
        else:
            print(f"[{now}] {message}")  # 初期画面では print にフォールバック

    def validate_first_screen(self):
        if not self.company_name.get().strip():
            messagebox.showwarning("入力エラー", "会社名は必須です")
            return False
        if not self.email.get().strip():
            messagebox.showwarning("入力エラー", "メールアドレスは必須です")
            return False
        if not self.password.get().strip():
            messagebox.showwarning("入力エラー", "パスワードは必須です")
            return False
        return True

    def validate_second_screen(self):
        from_str = self.date_from.get().strip()
        to_str = self.date_to.get().strip()

        if self.job1.get().strip() == "":
            messagebox.showwarning("入力エラー", "求人情報1は必須です")
            return False
        if self.pref1.get().strip() == "":
            messagebox.showwarning("入力エラー", "現住所1は必須です")
            return False
        if self.occ1.get().strip() == "":
            messagebox.showwarning("入力エラー", "経験職種1は必須です")
            return False

        try:
            age = int(self.max_age.get().strip())
            if age <= 0:
                raise ValueError
        except:
            messagebox.showwarning("入力エラー", "候補者年齢は正の整数で入力してください")
            return False

        if not from_str:
            messagebox.showwarning("入力エラー", "実行期間Fromは必須です")
            return False
        if not to_str:
            messagebox.showwarning("入力エラー", "実行期間Toは必須です")
            return False

        try:
            from_date = datetime.strptime(from_str, "%Y-%m-%d")
        except ValueError:
            messagebox.showwarning("入力エラー", f"実行期間Fromの形式または日付が不正です（例: 2025-05-08）")
            return False

        try:
            to_date = datetime.strptime(to_str, "%Y-%m-%d")
        except ValueError:
            messagebox.showwarning("入力エラー", f"実行期間Toの形式または日付が不正です（例: 2025-05-08）")
            return False

        if from_date > to_date:
            messagebox.showwarning("入力エラー", "実行期間FromはToより前の日付にしてください")
            return False

        return True

    def start_automation_thread(self):
        if not self.validate_second_screen():
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
                hours = int(wait_minutes) // 60
                minutes = int(wait_minutes) % 60
                self.log_window(f"開始時刻まで {hours}時間{minutes}分 待機します...")
                time.sleep(wait_minutes * 60)

            jobs = [self.job1.get(), self.job2.get(), self.job3.get()]
            prefs = [self.pref1.get(), self.pref2.get(), self.pref3.get()]
            occs  = [self.occ1.get(), self.occ2.get(), self.occ3.get()]

            loop_count = 0
            while datetime.now() <= to_time:
                for i in range(3):
                    job = jobs[i]
                    occ = occs[i]
                    for pref in prefs:
                        loop_count += 1
                        self.log_window(f"{loop_count}回目: 求人={job}, 現住所={pref}, 職種={occ} で処理開始")
                        self.run_condition_set(job, pref, occ)
                        self.log_window(f"{loop_count}回目: 処理完了")

                        # if datetime.now() + timedelta(minutes=20) > to_time:
                        #     return

                        # self.log_window("20分待機します...")
                        # time.sleep(20 * 60)
                        if datetime.now() + timedelta(minutes=1) > to_time:
                            return

                        self.log_window("1分待機します...")
                        time.sleep(1 * 60)
            self.log_window("実行期間終了。処理を終了します。")
        except Exception as e:
            self.log_window(f"処理中にエラー: {e}")
        self.log_window("処理を終了しました。")

    def run_condition_set(self, job_name, pref_name, occ_name):
        self.log_window(f"▼ 条件適用開始：求人={job_name}、現住所={pref_name}、職種={occ_name}")
        try:
            # 求人IDを取得
            job_id = self.job_title_id_map.get(job_name)
            if job_id:
                modal_button = self.wait.until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "div.switch.js_modalOpen"))
                )
                self.driver.execute_script("arguments[0].click();", modal_button)
                self.log_window("求人情報モーダルを開きました")
                time.sleep(1)

                job_elements = self.wait.until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a.js_optionWorkId"))
                )
                for i, job in enumerate(job_elements):
                    job_text = job.get_attribute("innerText").strip()
                    data_id = job.get_attribute("data-work_id")
                    self.log_window(f"取得求人候補[{i+1}]: {job_text}（ID: {data_id}）")
                    if data_id == job_id:
                        self.driver.execute_script("arguments[0].click();", job)
                        self.log_window(f"求人を選択: {job_text}")

                        try:
                            close_buttons = self.driver.find_elements(By.CLASS_NAME, "js_modalClose")
                            for btn in close_buttons:
                                try:
                                    self.driver.execute_script("arguments[0].click();", btn)
                                except:
                                    pass
                            time.sleep(1)  # モーダル閉じ待機
                        except Exception as e:
                            self.log_window(f"モーダルクローズ失敗: {e}")

                        break
            else:
                self.log_window(f"求人IDが見つかりませんでした: {job_name}")

            time.sleep(1)

            # 現住所選択
            if pref_name and pref_name != "選択しない":
                self.log_window(f"現住所で絞り込み: {pref_name}")
                try:
                    pref_code = self.pref_name_id_map.get(pref_name)
                    if pref_code:
                        select_element = self.wait.until(
                            EC.presence_of_element_located((By.ID, "md_select-candidatePrefecture"))
                        )
                        Select(select_element).select_by_value(pref_code)
                        self.log_window(f"現住所選択成功（{pref_name}: {pref_code}）")
                    else:
                        self.log_window(f"現住所コード取得失敗: {pref_name}")
                except Exception as e:
                    self.log_window(f"現住所選択エラー: {e}")
            else:
                self.log_window("現住所未指定のためスキップ")

            # 経験職種選択
            if occ_name and occ_name != "選択しない":
                self.log_window(f"経験職種で絞り込み: {occ_name}")
                try:
                    occ_code = self.occ_name_id_map.get(occ_name)
                    if occ_code:
                        select_element = self.wait.until(
                            EC.presence_of_element_located((By.ID, "md_select-candidateOccupation"))
                        )
                        Select(select_element).select_by_value(occ_code)
                        self.log_window(f"経験職種選択成功（{occ_name}: {occ_code}）")
                    else:
                        self.log_window(f"経験職種コード取得失敗: {occ_name}")
                except Exception as e:
                    self.log_window(f"経験職種選択エラー: {e}")
            else:
                self.log_window("経験職種未指定のためスキップ")

            # 絞り込みボタン押下
            refine_btn = self.driver.find_element(By.ID, "js_candidateRefinement")
            self.driver.execute_script("arguments[0].click();", refine_btn)
            time.sleep(2)

            # さらに読み込む最大まで
            show_more_count = 0
            for _ in range(9):
                try:
                    show_more = WebDriverWait(self.driver, 3).until(
                        EC.element_to_be_clickable((By.ID, "js_candidateShowMore"))
                    )
                    self.driver.execute_script("arguments[0].click();", show_more)
                    show_more_count += 1
                    time.sleep(2)
                except:
                    break

            self.log_window(f"『さらに読み込む』ボタンを {show_more_count} 回クリックしました")

            approach_count = 0
            error_count = 0

            # 候補者一覧から年齢条件に合う人を抽出して会ってみたい
            candidates = self.driver.find_elements(By.XPATH, '//div[@class="main"]')
            self.log_window(f"候補者一覧取得：{len(candidates)} 名")
            for idx, candidate in enumerate(candidates):
                self.log_window(f"[{idx+1}/{len(candidates)}] 候補者プロフィールを確認中...")
                try:
                    age_elements = candidate.find_elements(By.XPATH, './/span[contains(text(), "歳")]')
                    if not age_elements:
                        continue
                    age_text = age_elements[0].text
                    age = int(age_text.replace("歳", ""))
                    self.log_window(f"年齢: {age}")
                    if age > int(self.max_age.get()):
                        self.log_window("スキップ：年齢オーバー")
                        continue

                    profile_buttons = self.driver.find_elements(By.XPATH, '//a[contains(@class, "md_btn--detail")]')
                    if not profile_buttons:
                        self.log_window("プロフィールボタンが見つかりませんでした。スキップします。")
                        continue

                    # プロフィールを開く
                    self.driver.execute_script("arguments[0].click();", profile_buttons[0])
                    time.sleep(2)

                    # 「会ってみたいボタン」は詳細画面全体から探す
                    approach_btns = self.driver.find_elements(By.XPATH, '//a[contains(@class, "js_candidateApproach")]')
                    if approach_btns:
                    #    self.driver.execute_script("arguments[0].click();", approach_btns[0])
                        approach_count += 1
                        self.log_window("会ってみたいボタン押下")
                    else:
                        self.log_window("会ってみたいボタンなし")

                    # ポップアップ閉じる
                    body = self.driver.find_element(By.TAG_NAME, "body")
                    ActionChains(self.driver).move_to_element_with_offset(body, 10, 10).click().perform()
                    time.sleep(1)
                except Exception as e:
                    self.log_window(f"候補者処理エラー: {e}")
                    error_count += 1
            self.log_window(f"条件に合致した件数：{approach_count} 件")
            self.log_window(f"スキップまたはエラー件数: {error_count} 件")
        except Exception as e:
            self.log_window(f"run_condition_set エラー: {e}")

    def disable_inputs(self):
        for child in self.winfo_children():
            for widget in child.winfo_children():
                for sub_widget in widget.winfo_children():
                    try:
                        sub_widget.configure(state="disabled")
                    except:
                        pass
                try:
                    widget.configure(state="disabled")
                except:
                    pass

    def enable_inputs(self):
        for child in self.winfo_children():
            for widget in child.winfo_children():
                for sub_widget in widget.winfo_children():
                    try:
                        sub_widget.configure(state="normal")
                    except:
                        pass
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
                threading.Thread(target=self.driver.quit, daemon=True).start()
        except:
            pass
        self.destroy()

    def clear_screen(self):
        for widget in self.winfo_children():
            widget.destroy()

    def add_label_combobox_row(self, parent, label_prefix, variables, values, combo_list):
        row = ttk.Frame(parent)
        row.pack(anchor="w", pady=5)

        for i, var in enumerate(variables):
            label = ttk.Label(row, text=f"{label_prefix} {i+1}:", width=12, anchor="w")
            label.pack(side=tk.LEFT)

            combo = ttk.Combobox(row, textvariable=var, values=values, width=35)
            combo.pack(side=tk.LEFT, padx=10)

            combo_list.append(combo)

if __name__ == "__main__":
    app = EngageApp()
    app.mainloop()