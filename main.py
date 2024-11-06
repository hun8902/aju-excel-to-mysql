import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import mysql.connector
from mysql.connector import Error
import json
import requests
from datetime import datetime
import os
import threading

class ExcelToMySQLApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to MySQL Importer")
        self.root.geometry("600x700")
        
        # 스타일 설정
        style = ttk.Style()
        style.configure('TLabel', padding=5)
        style.configure('TEntry', padding=5)
        style.configure('TButton', padding=5)
        
        # 메인 프레임
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 파일 선택
        file_frame = ttk.LabelFrame(main_frame, text="Excel 파일 선택", padding="5")
        file_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.file_path = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file_path, width=50).pack(side=tk.LEFT, padx=5)
        ttk.Button(file_frame, text="찾아보기", command=self.browse_file).pack(side=tk.LEFT, padx=5)
        
        # 데이터베이스 설정
        db_frame = ttk.LabelFrame(main_frame, text="데이터베이스 설정", padding="5")
        db_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # DB 설정 입력 필드
        self.db_host = tk.StringVar(value="localhost")
        self.db_user = tk.StringVar(value="root")
        self.db_password = tk.StringVar()
        self.db_name = tk.StringVar(value="aju_erp")
        
        ttk.Label(db_frame, text="Host:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(db_frame, textvariable=self.db_host).grid(row=0, column=1, padx=5, pady=2, sticky=tk.EW)
        
        ttk.Label(db_frame, text="User:").grid(row=1, column=0, sticky=tk.W)
        ttk.Entry(db_frame, textvariable=self.db_user).grid(row=1, column=1, padx=5, pady=2, sticky=tk.EW)
        
        ttk.Label(db_frame, text="Password:").grid(row=2, column=0, sticky=tk.W)
        ttk.Entry(db_frame, textvariable=self.db_password, show="*").grid(row=2, column=1, padx=5, pady=2, sticky=tk.EW)
        
        ttk.Label(db_frame, text="Database:").grid(row=3, column=0, sticky=tk.W)
        ttk.Entry(db_frame, textvariable=self.db_name).grid(row=3, column=1, padx=5, pady=2, sticky=tk.EW)
        
        # 텔레그램 설정
        telegram_frame = ttk.LabelFrame(main_frame, text="텔레그램 설정", padding="5")
        telegram_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.telegram_token = tk.StringVar()
        self.telegram_chat_id = tk.StringVar()
        
        ttk.Label(telegram_frame, text="Bot Token:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(telegram_frame, textvariable=self.telegram_token).grid(row=0, column=1, padx=5, pady=2, sticky=tk.EW)
        
        ttk.Label(telegram_frame, text="Chat ID:").grid(row=1, column=0, sticky=tk.W)
        ttk.Entry(telegram_frame, textvariable=self.telegram_chat_id).grid(row=1, column=1, padx=5, pady=2, sticky=tk.EW)
        
        # 진행 상황 표시
        progress_frame = ttk.LabelFrame(main_frame, text="진행 상황", padding="5")
        progress_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.progress_text = tk.Text(progress_frame, height=15)
        self.progress_text.pack(fill=tk.BOTH, expand=True)
        
        # 스크롤바 추가
        scrollbar = ttk.Scrollbar(progress_frame, orient=tk.VERTICAL, command=self.progress_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.progress_text.configure(yscrollcommand=scrollbar.set)
        
        # 실행 버튼
        ttk.Button(main_frame, text="실행", command=self.start_import).pack(pady=10)
        
        # 설정 저장/불러오기
        self.load_settings()
        
    def browse_file(self):
        filename = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filename:
            self.file_path.set(filename)
    
    def log_progress(self, message):
        self.progress_text.insert(tk.END, f"{message}\n")
        self.progress_text.see(tk.END)
        self.root.update()
    
    def send_telegram_message(self, message):
        if self.telegram_token.get() and self.telegram_chat_id.get():
            try:
                url = f"https://api.telegram.org/bot{self.telegram_token.get()}/sendMessage"
                data = {
                    "chat_id": self.telegram_chat_id.get(),
                    "text": message
                }
                requests.post(url, json=data)
            except Exception as e:
                self.log_progress(f"텔레그램 메시지 전송 실패: {str(e)}")
    
    def save_settings(self):
        settings = {
            "db_host": self.db_host.get(),
            "db_user": self.db_user.get(),
            "db_name": self.db_name.get(),
            "telegram_token": self.telegram_token.get(),
            "telegram_chat_id": self.telegram_chat_id.get()
        }
        try:
            with open("settings.json", "w") as f:
                json.dump(settings, f)
        except Exception as e:
            self.log_progress(f"설정 저장 실패: {str(e)}")
    
    def load_settings(self):
        try:
            if os.path.exists("settings.json"):
                with open("settings.json", "r") as f:
                    settings = json.load(f)
                    self.db_host.set(settings.get("db_host", ""))
                    self.db_user.set(settings.get("db_user", ""))
                    self.db_name.set(settings.get("db_name", ""))
                    self.telegram_token.set(settings.get("telegram_token", ""))
                    self.telegram_chat_id.set(settings.get("telegram_chat_id", ""))
        except Exception as e:
            self.log_progress(f"설정 불러오기 실패: {str(e)}")
    
    def process_excel(self):
        try:
            # 결과 저장할 파일명 생성
            current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
            result_file = f"result_{current_time}.txt"
            
            total_results = []
            
            # 데이터베이스 연결 설정
            db_config = {
                'host': self.db_host.get(),
                'user': self.db_user.get(),
                'password': self.db_password.get(),
                'database': self.db_name.get()
            }
            
            # Excel 파일 처리
            excel_file = pd.ExcelFile(self.file_path.get())
            sheet_names = excel_file.sheet_names
            
            self.log_progress(f"처리할 시트 목록: {sheet_names}")
            total_inserted = 0
            
            with mysql.connector.connect(**db_config) as connection:
                for sheet_name in sheet_names:
                    try:
                        df = pd.read_excel(self.file_path.get(), sheet_name=sheet_name, engine='openpyxl', header=None)
                        
                        # 헤더 행 찾기
                        start_row = None
                        for idx, row in df.iterrows():
                            if '관리번호' in str(row.values):
                                start_row = idx
                                break
                        
                        if start_row is not None:
                            df = pd.read_excel(self.file_path.get(), sheet_name=sheet_name, engine='openpyxl', skiprows=start_row)
                            df.columns = [col.strip().replace(' ', '') for col in df.columns]
                            
                            cursor = connection.cursor()
                            inserted_count = 0
                            
                            for index, row in df.iterrows():
                                if pd.notna(row['관리번호']):
                                    insert_query = """
                                        INSERT INTO aju_facilities 
                                        (fc_code, fc_purpose, fc_use, fc_name, fc_size, fc_model, fc_maker, fc_buy_date)
                                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                                    """
                                    
                                    data = (
                                        str(row['관리번호']),
                                        str(row['사용부서']),
                                        str(row['용도']),
                                        str(row['설비명칭']),
                                        str(row['규격/용량']),
                                        str(row['모델명']),
                                        str(row['제작사']),
                                        str(row['구입일자'])
                                    )
                                    
                                    cursor.execute(insert_query, data)
                                    inserted_count += 1
                                    
                                    if inserted_count % 100 == 0:
                                        connection.commit()
                                        self.log_progress(f"{sheet_name}: {inserted_count}개 처리됨")
                            
                            connection.commit()
                            cursor.close()
                            
                            result = f"{sheet_name} 시트: {inserted_count}개 데이터 삽입 완료"
                            total_results.append(result)
                            self.log_progress(result)
                            total_inserted += inserted_count
                            
                        else:
                            result = f"{sheet_name} 시트: '관리번호' 열을 찾을 수 없음"
                            total_results.append(result)
                            self.log_progress(result)
                            
                    except Exception as e:
                        result = f"{sheet_name} 시트 처리 중 오류: {str(e)}"
                        total_results.append(result)
                        self.log_progress(result)
            
            # 결과 파일 저장
            with open(result_file, "w", encoding="utf-8") as f:
                f.write("\n".join(total_results))
            
            # 텔레그램 메시지 전송
            summary = f"엑셀 데이터 처리 완료\n총 입력 건수: {total_inserted}\n자세한 내용은 {result_file} 파일을 확인하세요."
            self.send_telegram_message(summary)
            
            messagebox.showinfo("완료", f"처리가 완료되었습니다.\n결과는 {result_file} 파일에서 확인할 수 있습니다.")
            
        except Exception as e:
            error_message = f"오류 발생: {str(e)}"
            self.log_progress(error_message)
            self.send_telegram_message(error_message)
            messagebox.showerror("오류", error_message)
    
    def start_import(self):
        if not self.file_path.get():
            messagebox.showerror("오류", "Excel 파일을 선택해주세요.")
            return
        
        # 설정 저장
        self.save_settings()
        
        # 별도 스레드에서 처리 시작
        self.progress_text.delete(1.0, tk.END)
        self.log_progress("처리 시작...")
        threading.Thread(target=self.process_excel, daemon=True).start()

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelToMySQLApp(root)
    root.mainloop()