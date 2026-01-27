import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import requests
import threading
import os
import sys
from pykakasi import kakasi

# --- Configuration ---
# API endpoints
GEOCODING_API_ENDPOINT = "https://maps.googleapis.com/maps/api/geocode/json"
ELEVATION_API_ENDPOINT = "https://maps.googleapis.com/maps/api/elevation/json"

# --- Data for Island Mapping ---
ISLAND_MAP = {
    '北海道': 'Hokkaido',
    '青森県': 'Honshu', '岩手県': 'Honshu', '宮城県': 'Honshu', '秋田県': 'Honshu', '山形県': 'Honshu', '福島県': 'Honshu',
    '茨城県': 'Honshu', '栃木県': 'Honshu', '群馬県': 'Honshu', '埼玉県': 'Honshu', '千葉県': 'Honshu', '東京都': 'Honshu', '神奈川県': 'Honshu',
    '新潟県': 'Honshu', '富山県': 'Honshu', '石川県': 'Honshu', '福井県': 'Honshu', '山梨県': 'Honshu', '長野県': 'Honshu', '岐阜県': 'Honshu', '静岡県': 'Honshu', '愛知県': 'Honshu',
    '三重県': 'Honshu', '滋賀県': 'Honshu', '京都府': 'Honshu', '大阪府': 'Honshu', '兵庫県': 'Honshu', '奈良県': 'Honshu', '和歌山県': 'Honshu',
    '鳥取県': 'Honshu', '島根県': 'Honshu', '岡山県': 'Honshu', '広島県': 'Honshu', '山口県': 'Honshu',
    '徳島県': 'Shikoku', '香川県': 'Shikoku', '愛媛県': 'Shikoku', '高知県': 'Shikoku',
    '福岡県': 'Kyushu', '佐賀県': 'Kyushu', '長崎県': 'Kyushu', '熊本県': 'Kyushu', '大分県': 'Kyushu', '宮崎県': 'Kyushu', '鹿児島県': 'Kyushu',
    '沖縄県': 'Okinawa Islands'
}

# --- Kakasi setup ---
kks = kakasi()
kks.setMode("H", "a")
kks.setMode("K", "a")
kks.setMode("J", "a")
conv = kks.getConverter()

class LabelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("標本ラベルデータ生成ツール")
        self.root.geometry("600x550")

        # --- Variables ---
        self.api_key_var = tk.StringVar()
        self.input_file_path = tk.StringVar()
        self.status_var = tk.StringVar(value="待機中")
        
        # --- UI Layout ---
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 1. API Key Section
        api_frame = ttk.LabelFrame(main_frame, text="1. Google Maps API設定", padding="10")
        api_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(api_frame, text="APIキー:").pack(side=tk.LEFT)
        self.api_entry = ttk.Entry(api_frame, textvariable=self.api_key_var, width=40, show="*")
        self.api_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        # Default value for convenience (remove in production distribution)
        self.api_key_var.set("ここにあなたのAPIキーを入力")

        # 2. File Selection Section
        file_frame = ttk.LabelFrame(main_frame, text="2. データファイルの選択 (CSV / Excel)", padding="10")
        file_frame.pack(fill=tk.X, pady=5)
        
        ttk.Entry(file_frame, textvariable=self.input_file_path, state="readonly").pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(file_frame, text="参照...", command=self.browse_file).pack(side=tk.LEFT, padx=5)

        # 3. Options (Column Mapping)
        opt_frame = ttk.LabelFrame(main_frame, text="3. 列名の設定 (CSVの列名と一致させてください)", padding="10")
        opt_frame.pack(fill=tk.X, pady=5)

        self.col_entries = {}
        cols = [
            ("緯度の列名", "latitude"), ("経度の列名", "longitude"),
            ("日付の列名", "採集年月日"), ("採集方法の列名", "採集方法"),
            ("採集者名の列名", "採集者名")
        ]
        
        for i, (label, default) in enumerate(cols):
            row = i // 2
            col = (i % 2) * 2
            ttk.Label(opt_frame, text=f"{label}:").grid(row=row, column=col, sticky="e", padx=5, pady=2)
            entry = ttk.Entry(opt_frame, width=15)
            entry.insert(0, default)
            entry.grid(row=row, column=col+1, sticky="w", padx=5, pady=2)
            self.col_entries[label] = entry

        # 4. Execution Section
        run_frame = ttk.Frame(main_frame, padding="10")
        run_frame.pack(fill=tk.X, pady=10)
        
        self.run_btn = ttk.Button(run_frame, text="処理開始", command=self.start_process)
        self.run_btn.pack(fill=tk.X, ipady=5)
        
        self.progress = ttk.Progressbar(run_frame, mode='determinate')
        self.progress.pack(fill=tk.X, pady=5)
        
        ttk.Label(run_frame, textvariable=self.status_var).pack()

    def browse_file(self):
        filetypes = (("CSV files", "*.csv"), ("Excel files", "*.xlsx;*.xls"), ("All files", "*.*"))
        filename = filedialog.askopenfilename(title="ファイルを開く", filetypes=filetypes)
        if filename:
            self.input_file_path.set(filename)

    def start_process(self):
        api_key = self.api_key_var.get()
        input_path = self.input_file_path.get()
        
        if not input_path:
            messagebox.showwarning("警告", "入力ファイルを選択してください。")
            return
        if not api_key:
            messagebox.showwarning("警告", "APIキーを入力してください。")
            return

        # Disable button
        self.run_btn.config(state="disabled")
        self.progress['value'] = 0
        self.status_var.set("処理中...")

        # Run in a separate thread to keep UI responsive
        thread = threading.Thread(target=self.process_data, args=(api_key, input_path))
        thread.start()

    def process_data(self, api_key, input_path):
        try:
            # Read Data
            if input_path.endswith('.csv'):
                df = pd.read_csv(input_path)
            else:
                df = pd.read_excel(input_path)
            
            col_map = {k: v.get() for k, v in self.col_entries.items()}
            total_rows = len(df)
            results = []

            for index, row in df.iterrows():
                lat = row.get(col_map["緯度の列名"])
                lon = row.get(col_map["経度の列名"])
                
                if pd.notna(lat) and pd.notna(lon):
                    addr_info = self.get_google_address(lat, lon, api_key)
                    elev = self.get_elevation(lat, lon, api_key)
                    if elev is not None:
                        addr_info['alt'] = elev
                    results.append(addr_info)
                else:
                    results.append({'status': 'データなし'})
                
                # Update progress
                progress_val = (index + 1) / total_rows * 100
                self.root.after(0, lambda v=progress_val: self.progress.configure(value=v))
                self.root.after(0, lambda i=index+1, t=total_rows: self.status_var.set(f"処理中: {i}/{t} 件"))

            # Combine results
            results_df = pd.DataFrame(results)
            df_combined = pd.concat([df.reset_index(drop=True), results_df], axis=1)

            # Generate Label Text
            df_combined['label'] = df_combined.apply(
                lambda r: self.create_label_text(r, col_map), axis=1
            )

            # Organize Columns
            # Basic address info columns
            new_cols = ['地点名の表記', '国名', '県名', '地点(ローマ字)', '島・大陸名', '市区町村', '市区町村種別', 'alt', 'status', 'label']
            # Keep original columns + new columns
            final_cols = df.columns.tolist() + [c for c in new_cols if c not in df.columns]
            df_output = df_combined.reindex(columns=final_cols)

            # Save Output
            output_path = os.path.splitext(input_path)[0] + "_labeled.xlsx"
            df_output.to_excel(output_path, index=False)

            self.root.after(0, lambda: messagebox.showinfo("完了", f"処理が完了しました！\n\n保存先:\n{output_path}"))
            self.root.after(0, lambda: self.status_var.set("完了"))

        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("エラー", f"予期せぬエラーが発生しました:\n{e}"))
            self.root.after(0, lambda: self.status_var.set("エラー発生"))
        
        finally:
            self.root.after(0, lambda: self.run_btn.config(state="normal"))

    # --- Logic Functions (Same as before, adapted for Class) ---

    def get_elevation(self, lat, lon, api_key):
        params = {'locations': f'{lat},{lon}', 'key': api_key}
        try:
            res = requests.get(ELEVATION_API_ENDPOINT, params=params, timeout=5).json()
            if res['status'] == 'OK': return int(round(res['results'][0]['elevation']))
        except: return None
        return None

    def get_google_address(self, lat, lon, api_key):
        # Initialize structure
        res_data = {k: '' for k in ['地点名の表記', '国名', '県名', '地点(ローマ字)', '島・大陸名', '市区町村', '市区町村種別', 'alt']}
        res_data['status'] = 'エラー'

        params = {
            'latlng': f'{lat},{lon}', 'key': api_key, 'language': 'ja',
            'result_type': 'political|locality|sublocality|neighborhood|premise|subpremise'
        }
        try:
            resp = requests.get(GEOCODING_API_ENDPOINT, params=params, timeout=5).json()
            
            if resp['status'] == 'OK':
                first = resp['results'][0]
                
                # Full Address (JP)
                full = first.get('formatted_address', '')
                full = full.replace('日本、', '', 1)
                if ' ' in full and full[:full.find(' ')].replace('〒','').replace('-','').isdigit():
                    full = full[full.find(' ')+1:]
                res_data['地点名の表記'] = full

                # Components
                comp_map = {'country': '', 'administrative_area_level_1': '', 'locality': '', 'sublocality_level_1': ''}
                for c in first.get('address_components', []):
                    types = c.get('types', [])
                    for t in comp_map:
                        if t in types:
                            if t == 'country': comp_map[t] = c.get('short_name')
                            else: comp_map[t] = c.get('long_name')
                
                res_data['国名'] = comp_map['country']
                res_data['県名'] = comp_map['administrative_area_level_1']
                
                # Locality logic
                muni = comp_map['locality']
                point_jp = comp_map['sublocality_level_1']
                if not muni: muni = comp_map['administrative_area_level_1'] # Fallback

                res_data['市区町村'] = muni
                
                # Island Logic
                island = ''
                if muni and muni in ISLAND_MAP.values(): island = muni # If muni is known island name
                elif res_data['県名'] in ISLAND_MAP: island = ISLAND_MAP[res_data['県名']]
                res_data['島・大陸名'] = island

                # Romaji Conversion
                target_name = point_jp if point_jp else muni
                res_data['地点(ローマ字)'] = conv.do(target_name)
                
                # Suffixes for Japan
                if res_data['国名'] == 'JP':
                    if res_data['県名']: res_data['県名'] += '-pref'
                    if muni:
                         if muni.endswith('市'): res_data['市区町村'] += '-shi'
                         elif muni.endswith('区'): res_data['市区町村'] += '-ku'
                         elif muni.endswith('町'): res_data['市区町村'] += '-cho'
                         elif muni.endswith('村'): res_data['市区町村'] += '-mura'

                res_data['status'] = '成功'
            
            else:
                res_data['status'] = f"APIエラー: {resp.get('status')}"

        except Exception as e:
            res_data['status'] = f"通信エラー: {e}"
        
        return res_data

    def create_label_text(self, row, col_map):
        # Extract basic data
        date = str(row.get(col_map["日付の列名"], ''))
        method = str(row.get(col_map["採集方法の列名"], ''))
        collector = str(row.get(col_map["採集者名の列名"], ''))
        
        # Extract API data
        addr = row.get('地点名の表記', '')
        elev = row.get('alt', '')
        lat = row.get(col_map["緯度の列名"], '')
        lon = row.get(col_map["経度の列名"], '')

        # Build Text
        lines = []
        lines.append(f"JAPAN: {addr}")
        lines.append(f"GPS({elev}m)")
        
        line3 = []
        if date and date != 'nan': line3.append(date)
        if method and method != 'nan': line3.append(method)
        if collector and collector != 'nan': line3.append(collector)
        if line3: lines.append(". ".join(line3) + ".")
        
        lines.append(f"N{lat}, E{lon}")
        
        return "\n".join(lines)

if __name__ == "__main__":
    root = tk.Tk()
    app = LabelApp(root)
    root.mainloop()