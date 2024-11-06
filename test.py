import pandas as pd
import mysql.connector
from mysql.connector import Error

# 엑셀 파일 경로
excel_file_path = './test1.xlsx'

# MySQL 데이터베이스 연결 설정
db_config = {
    'host': 'localhost',
    'user': 'root',
    'password': '',
    'database': 'aju_erp'
}

def find_header_row(df):
    """'관리번호' 열이 있는 행을 찾는 함수"""
    for idx, row in df.iterrows():
        if '관리번호' in str(row.values):
            return idx
    return None

def clean_column_name(column_name):
    """열 이름에서 공백을 제거하고 표준화하는 함수"""
    return column_name.strip().replace(' ', '')

def process_sheet(sheet_name, connection):
    """개별 시트 처리 함수"""
    try:
        # 먼저 시트를 읽어서 헤더 행 찾기
        df = pd.read_excel(excel_file_path, sheet_name=sheet_name, engine='openpyxl', header=None)
        
        start_row = find_header_row(df)
        if start_row is None:
            print(f"{sheet_name} 시트에서 '관리번호' 열을 찾을 수 없습니다.")
            return
        
        # 데이터 다시 읽기 (찾은 시작 행부터)
        df = pd.read_excel(excel_file_path, sheet_name=sheet_name, engine='openpyxl', skiprows=start_row)
        print(f"\n{sheet_name} 시트 처리 시작...")
        print(f"원본 데이터 열 이름:", df.columns.tolist())
        
        # 열 이름 정리 (공백 제거 및 표준화)
        df.columns = [clean_column_name(col) for col in df.columns]
        print(f"정리된 데이터 열 이름:", df.columns.tolist())
        
        # NaN 값을 None으로 변환
        df = df.where(pd.notnull(df), None)
        
        cursor = connection.cursor()
        
        # 각 행을 순회하며 데이터 삽입
        inserted_count = 0
        for index, row in df.iterrows():
            if pd.notna(row['관리번호']):  # 관리번호가 있는 행만 처리
                # 데이터 삽입 SQL 쿼리
                insert_query = """
                    INSERT INTO aju_facilities 
                    (fc_code, fc_purpose, fc_use, fc_name, fc_size, fc_model, fc_maker, fc_buy_date)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                """
                
                # 데이터 튜플 생성
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
                
                # 데이터 확인 출력 (디버깅용)
                print(f"데이터 삽입: {data}")
                
                # 데이터 삽입 실행
                cursor.execute(insert_query, data)
                inserted_count += 1
                
                # 매 100건마다 커밋
                if inserted_count % 100 == 0:
                    connection.commit()
        
        # 남은 변경사항 커밋
        connection.commit()
        print(f"{sheet_name} 시트에서 {inserted_count}개의 데이터가 성공적으로 삽입되었습니다.")
        cursor.close()
        
    except Exception as e:
        print(f"{sheet_name} 시트 처리 중 오류 발생:", e)
        print("현재 처리 중인 데이터:", row.to_dict() if 'row' in locals() else "데이터 없음")
        import traceback
        print(traceback.format_exc())

try:
    # MySQL 데이터베이스 연결
    connection = mysql.connector.connect(**db_config)
    if connection.is_connected():
        print("MySQL 데이터베이스에 연결되었습니다.")
        
        # 모든 시트 이름 가져오기
        excel_file = pd.ExcelFile(excel_file_path)
        sheet_names = excel_file.sheet_names
        
        print(f"처리할 시트 목록: {sheet_names}")
        
        # 각 시트 순서대로 처리
        for sheet_name in sheet_names:
            process_sheet(sheet_name, connection)

except Error as e:
    print("MySQL 연결 오류:", e)
    
finally:
    if 'connection' in locals() and connection.is_connected():
        connection.close()
        print("\nMySQL 연결이 닫혔습니다.")