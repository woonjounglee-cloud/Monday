# Excel Scraper and Merger

특정 웹사이트에서 엑셀 파일을 다운로드하고 데이터를 가공하여 하나의 파일로 병합하는 Python 프로그램입니다.

## 기능

- 여러 URL에서 엑셀 파일 자동 다운로드
- 다운로드한 엑셀 파일 데이터 가공 및 정리
- 여러 엑셀 파일을 하나의 파일로 병합
- 각 데이터의 출처(Source) 추적
- 로깅을 통한 프로세스 모니터링

## 설치 방법

### 1. Python 환경 확인

Python 3.8 이상이 설치되어 있어야 합니다.

```bash
python --version
```

### 2. 필요한 패키지 설치

```bash
pip install -r requirements.txt
```

## 사용 방법

### 1. URL 설정

`excel_scraper_merger.py` 파일을 열어서 다운로드할 파일의 URL을 설정합니다:

```python
def main():
    scraper = ExcelScraperMerger()

    # 실제 URL로 변경하세요
    scraper.urls['Field_Protocol_Driving_Test'] = 'https://your-url.com/file1.xlsx'
    scraper.urls['Field_Protocol_Stationary_Call_Test'] = 'https://your-url.com/file2.xlsx'
    scraper.urls['Field_Protocol_VQ_Test'] = 'https://your-url.com/file3.xlsx'

    # 실행
    result = scraper.run(merge=True)
```

### 2. 프로그램 실행

```bash
python excel_scraper_merger.py
```

### 3. 실행 옵션

#### 모든 파일을 하나로 병합
```python
result = scraper.run(merge=True)
```

#### 각 파일을 개별적으로 처리
```python
result = scraper.run(merge=False)
```

## 프로그램 구조

```
Monday/
├── excel_scraper_merger.py   # 메인 스크립트
├── requirements.txt           # 필요한 패키지 목록
├── downloads/                 # 다운로드된 파일 저장 폴더 (자동 생성)
└── output/                    # 처리된 결과 파일 저장 폴더 (자동 생성)
```

## 주요 클래스 및 메서드

### ExcelScraperMerger 클래스

#### 주요 메서드

- `download_excel_file(url, filename)`: URL에서 엑셀 파일 다운로드
- `process_excel_file(filepath, sheet_name)`: 엑셀 파일 읽기 및 처리
- `add_source_column(df, source_name)`: 출처 컬럼 추가
- `download_all_files()`: 모든 URL에서 파일 다운로드
- `process_all_files(downloaded_files)`: 다운로드된 모든 파일 처리
- `merge_dataframes(dataframes)`: 여러 데이터프레임 병합
- `save_to_excel(df, filename)`: 엑셀 파일로 저장
- `run(merge)`: 전체 프로세스 실행

## 코드 사용 예제

### 기본 사용

```python
from excel_scraper_merger import ExcelScraperMerger

# 인스턴스 생성
scraper = ExcelScraperMerger()

# URL 설정
scraper.urls['Field_Protocol_Driving_Test'] = 'https://example.com/file1.xlsx'
scraper.urls['Field_Protocol_Stationary_Call_Test'] = 'https://example.com/file2.xlsx'
scraper.urls['Field_Protocol_VQ_Test'] = 'https://example.com/file3.xlsx'

# 실행
result = scraper.run(merge=True)
print(f"결과 파일: {result}")
```

### 개별 파일 처리

```python
scraper = ExcelScraperMerger()
scraper.urls['Field_Protocol_Driving_Test'] = 'https://example.com/file1.xlsx'

# 파일 다운로드
filepath = scraper.download_excel_file(
    scraper.urls['Field_Protocol_Driving_Test'],
    'driving_test.xlsx'
)

# 파일 처리
df = scraper.process_excel_file(filepath)

# 출처 추가
df = scraper.add_source_column(df, 'Field_Protocol_Driving_Test')

# 저장
scraper.save_to_excel(df, 'processed_driving_test.xlsx')
```

### 커스텀 데이터 가공

```python
scraper = ExcelScraperMerger()

# 파일 다운로드 및 처리
downloaded_files = scraper.download_all_files()
processed_data = []

for name, filepath in downloaded_files.items():
    df = scraper.process_excel_file(filepath)

    # 커스텀 데이터 가공 예시
    # 특정 컬럼만 선택
    # df = df[['컬럼1', '컬럼2', '컬럼3']]

    # 조건에 맞는 행만 필터링
    # df = df[df['컬럼1'] > 100]

    # 새로운 컬럼 추가
    # df['새컬럼'] = df['컬럼1'] * 2

    df = scraper.add_source_column(df, name)
    processed_data.append(df)

# 병합 및 저장
merged_df = scraper.merge_dataframes(processed_data)
scraper.save_to_excel(merged_df, 'custom_merged.xlsx')
```

## 출력 파일

### 병합 모드 (merge=True)
- `output/merged_field_protocols.xlsx`: 모든 파일이 병합된 결과 파일
- 각 행의 `Source` 컬럼에서 데이터의 출처 확인 가능

### 개별 모드 (merge=False)
- `output/Field_Protocol_Driving_Test_processed.xlsx`
- `output/Field_Protocol_Stationary_Call_Test_processed.xlsx`
- `output/Field_Protocol_VQ_Test_processed.xlsx`

## 로그

프로그램 실행 중 다음과 같은 로그가 출력됩니다:

```
2025-11-17 10:00:00 - INFO - === Excel Scraper and Merger 시작 ===
2025-11-17 10:00:01 - INFO - 1. 파일 다운로드 중...
2025-11-17 10:00:02 - INFO - 다운로드 시작: Field_Protocol_Driving_Test.xlsx
2025-11-17 10:00:05 - INFO - 다운로드 완료: downloads/Field_Protocol_Driving_Test.xlsx
...
2025-11-17 10:00:20 - INFO - === 완료 ===
```

## 에러 처리

프로그램은 다음과 같은 상황을 처리합니다:

- URL이 설정되지 않은 경우
- 파일 다운로드 실패
- 파일이 존재하지 않는 경우
- 엑셀 파일 읽기 실패
- 병합 실패
- 파일 저장 실패

모든 에러는 로그로 기록되며, 프로그램은 가능한 파일만 처리합니다.

## 주의사항

1. URL은 반드시 실제 다운로드 가능한 엑셀 파일 주소여야 합니다
2. 파일 크기가 큰 경우 다운로드 시간이 오래 걸릴 수 있습니다
3. 인터넷 연결이 필요합니다
4. 엑셀 파일 형식은 .xlsx 또는 .xls를 지원합니다

## 라이센스

MIT License

## 문의

문제가 발생하거나 개선 사항이 있으면 이슈를 등록해주세요.
