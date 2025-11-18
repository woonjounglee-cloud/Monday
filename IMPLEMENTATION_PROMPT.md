# Monday 웹서버 구현 프롬프트

## 프로젝트 개요
Flask 기반 웹 애플리케이션으로, Excel 파일 처리 및 모델 데이터베이스 관리 시스템을 구현합니다.

---

## 1. 프로젝트 구조

```
Monday/
├── app.py                      # Flask 메인 서버
├── excel_processor.py          # Excel 파일 처리 로직
├── excel_scraper_merger.py     # Excel 스크래퍼 (선택적)
├── requirements.txt            # 필요한 패키지 목록
├── templates/
│   └── index.html             # 웹 인터페이스
├── uploads/                   # 업로드된 파일 저장 폴더 (자동 생성)
├── output/                    # 결과 파일 저장 폴더 (자동 생성)
├── db/                        # 데이터베이스 폴더 (자동 생성)
│   └── model_manager.xlsx     # 모델담당자 데이터베이스
└── drivers/                   # Chrome WebDriver 폴더 (선택적)
```

---

## 2. requirements.txt 작성

다음 패키지들을 requirements.txt에 포함:

```
pandas>=2.0.0
openpyxl>=3.1.0
xlrd>=2.0.1
requests>=2.31.0
flask>=3.0.0
selenium>=4.0.0
webdriver-manager>=4.0.0
```

---

## 3. excel_processor.py 구현

### 3.1 클래스 구조

`ExcelProcessor` 클래스를 생성하고 다음 속성과 메서드를 구현:

```python
#!/usr/bin/env python3
"""
Excel Processor for FA Task Schedule
Schedule 파일과 모델담당자 데이터베이스를 병합하여 결과 파일 생성
"""

import pandas as pd
from datetime import datetime
from typing import Dict, Optional
import logging
import os

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class ExcelProcessor:
    """엑셀 파일 처리 및 병합 클래스"""

    # 추출할 컬럼 목록
    REQUIRED_COLUMNS = [
        '검증항목', '과제명', '개발모델명', '검증단계',
        'PRA', '의뢰일', '완료요청일', '검증 PL'
    ]

    # 컬럼 별칭 (여러 이름으로 불릴 수 있는 컬럼)
    COLUMN_ALIASES = {
        '의뢰일': ['의뢰일', '외뢰일', '의뢰', '외뢰'],
        '검증 PL': ['검증 PL', '검증PL', 'PL'],
    }

    # 검증항목 우선순위 (낮은 숫자가 우선)
    VERIFICATION_ITEM_PRIORITY = {
        '주행': 1,      # 필드 프로토콜_주행 시험 / 필드 프로토콜_주행시험
        '고정점': 2,    # 필드 프로토콜_고정점 송수신 시험
        '송수화': 3     # 필드 프로토콜_송수화 시험
    }

    def __init__(self):
        """초기화"""
        self.schedule_df = None
        self.model_manager_df = None
```

#### 인스턴스 속성
- `self.schedule_df`: Schedule 데이터프레임
- `self.model_manager_df`: 모델담당자 데이터프레임

### 3.2 핵심 메서드 구현

#### `get_week_number() -> str`
현재 날짜의 주차를 계산하여 W{주차} 형식으로 반환

```python
def get_week_number(self) -> str:
    """
    현재 날짜의 주차를 계산하여 W{주차} 형식으로 반환

    Returns:
        str: W{주차} 형식 (예: W47)
    """
    now = datetime.now()
    week_number = now.isocalendar()[1]
    return f"W{week_number:02d}"
```

#### `get_engine_for_file(filepath: str) -> str`
파일 확장자에 따라 적절한 pandas 엔진 선택

```python
def get_engine_for_file(self, filepath: str) -> str:
    """
    파일 확장자에 따라 적절한 pandas 엔진 선택

    Args:
        filepath: 파일 경로

    Returns:
        str: 엔진 이름
    """
    ext = os.path.splitext(filepath)[1].lower()
    if ext == '.xls':
        return 'xlrd'
    elif ext == '.xlsx':
        return 'openpyxl'
    else:
        return 'openpyxl'  # 기본값
```

#### `read_excel_with_fallback(filepath: str, **kwargs) -> pd.DataFrame`
여러 엔진을 시도하여 엑셀 파일 읽기

```python
def read_excel_with_fallback(self, filepath: str, **kwargs) -> pd.DataFrame:
    """
    여러 엔진을 시도하여 엑셀 파일 읽기

    Args:
        filepath: 파일 경로
        **kwargs: pandas read_excel 추가 인자

    Returns:
        pd.DataFrame: 읽은 데이터프레임
    """
    # 파일 확장자에 맞는 엔진 선택
    primary_engine = self.get_engine_for_file(filepath)
    engines = [primary_engine]

    # fallback 엔진 추가
    if primary_engine == 'openpyxl':
        engines.append('xlrd')
    else:
        engines.append('openpyxl')

    last_error = None

    for engine in engines:
        try:
            logger.info(f"엔진 '{engine}'으로 파일 읽기 시도: {filepath}")
            df = pd.read_excel(filepath, engine=engine, **kwargs)
            logger.info(f"성공: 엔진 '{engine}'으로 파일 읽기 완료")
            return df
        except Exception as e:
            logger.warning(f"엔진 '{engine}' 실패: {e}")
            last_error = e
            continue

    # 모든 엔진 실패 시
    raise Exception(f"모든 엔진으로 파일 읽기 실패. 마지막 에러: {last_error}")
```

#### `find_header_row(df: pd.DataFrame, required_columns: list) -> int`
첫 10개 행에서 헤더 위치 찾기

```python
def find_header_row(self, df: pd.DataFrame, required_columns: list) -> int:
    """
    헤더 행의 위치를 찾음

    Args:
        df: 데이터프레임
        required_columns: 찾을 컬럼 목록

    Returns:
        int: 헤더 행 번호 (0-based), 찾지 못하면 0
    """
    # 첫 10개 행에서 헤더 찾기
    for row_idx in range(min(10, len(df))):
        row_values = df.iloc[row_idx].astype(str).tolist()

        # 필요한 컬럼 중 50% 이상이 있으면 헤더로 판단
        matches = sum(1 for col in required_columns if any(col in str(val) for val in row_values))
        if matches >= len(required_columns) * 0.5:
            logger.info(f"헤더 행 발견: {row_idx}번째 행")
            return row_idx

    logger.warning("헤더 행을 찾지 못함. 첫 번째 행을 헤더로 사용")
    return 0
```

#### `read_schedule_file(filepath: str) -> pd.DataFrame`
1. 헤더 위치 자동 감지
2. 필요한 컬럼 추출 (부분 일치 및 별칭 허용)
3. 빈 행 제거
4. 누락된 컬럼은 빈 문자열로 추가

#### `read_model_manager_file(filepath: str) -> pd.DataFrame`
1. 모델담당자 파일 읽기
2. 필요한 컬럼: '개발모델명', '모델담당자', 'AP/CP'
3. 컬럼명 매핑 (부분 일치 허용)

#### `get_verification_priority(verification_item: str) -> int`
- 검증항목의 우선순위 반환
- '주행' → 1, '고정점'/'송수신' → 2, '송수화' → 3
- 정의되지 않은 항목 → 999

#### `sort_schedule_data(df: pd.DataFrame) -> pd.DataFrame`
1. 검증항목 우선순위 계산
2. 정렬 순서: 검증항목 우선순위 → PRA 오름차순
3. PRA가 없는 행은 맨 뒤로 이동

#### `vlookup_model_manager(schedule_df: pd.DataFrame) -> pd.DataFrame`
1. 개발모델명을 기준으로 모델담당자 정보 매칭
2. pandas의 merge 함수 사용
3. **특별 규칙**: 송수화 시험은 모델담당자를 무조건 '이운정'으로 설정
4. 매칭되지 않는 항목은 빈 문자열

#### `save_result(df: pd.DataFrame, output_path: str) -> bool`
1. 최종 컬럼 순서: REQUIRED_COLUMNS + ['모델담당자', 'AP/CP']
2. openpyxl로 저장
3. 모든 셀에 가운데 정렬 적용

#### `process(schedule_file: str, model_manager_file: str, output_path: str) -> bool`
전체 프로세스 실행:
1. Schedule 파일 읽기
2. 모델담당자 데이터베이스 읽기
3. 데이터 정렬
4. VLOOKUP으로 모델담당자 정보 매칭
5. 결과 저장

---

## 4. app.py 구현

### 4.1 Flask 앱 설정

```python
app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB 제한
app.config['SECRET_KEY'] = 'your-secret-key-here'

# 업로드 폴더 설정
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'
DB_FOLDER = 'db'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

# 폴더 생성
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(DB_FOLDER, exist_ok=True)

# 모델담당자 데이터베이스 파일 경로
MODEL_DB_PATH = os.path.join(DB_FOLDER, 'model_manager.xlsx')

# 검증허브 URL (환경 변수로 설정 가능)
TESTHUB_URL = os.getenv('TESTHUB_URL', 'https://example.com')

# 검증허브 처리 결과 저장용 전역 변수
last_testhub_result = None
```

### 4.2 유틸리티 함수

#### `allowed_file(filename) -> bool`
- 파일 확장자가 허용된 확장자인지 확인

#### `load_model_database() -> pd.DataFrame`
- 모델담당자 데이터베이스 로드
- 컬럼 순서: ['과제명', '개발모델명', '검증PL', '모델담당자', 'AP/CP']
- 파일이 없으면 빈 데이터프레임 생성

#### `save_model_database(df: pd.DataFrame) -> bool`
- 모델담당자 데이터베이스 저장
- 컬럼 순서 재정렬

### 4.3 라우트 구현

#### `GET /`
- 메인 페이지 (index.html 렌더링)

#### `POST /upload`
Schedule 파일 업로드 및 처리:
1. 파일 유효성 검사
2. ExcelProcessor로 처리
3. 주차 계산하여 파일명 생성 (예: W47_FA_과제일정.xlsx)
4. 결과 파일 저장
5. 성공 시 download_url과 preview_url 반환

#### `GET /preview/<filename>`
- 결과 파일 미리보기
- pandas로 엑셀 읽어서 JSON 형식으로 반환
- 형식: `{columns: [...], rows: [[...]], total_rows: N}`

#### `GET /download/<path:filename>`
- 결과 파일 다운로드
- 경로 검증으로 디렉토리 탈출 방지
- 한글 파일명 지원

### 4.4 Model Info API 구현

#### `GET /api/model-info/all`
- 전체 모델 정보 조회

#### `POST /api/model-info/search`
- 키워드로 모델 정보 검색
- 과제명 또는 개발모델명에서 검색

#### `POST /api/model-info/add`
- 새 모델 정보 추가
- 필수 필드: 과제명, 개발모델명, 검증PL, 모델담당자, AP/CP

#### `POST /api/model-info/update`
- 모델 정보 수정
- 요청: row_index, field, value

#### `POST /api/model-info/delete`
- 모델 정보 삭제
- 요청: row_index

#### `GET /api/model-info/export`
- 모델담당자 데이터베이스를 Excel 파일로 내보내기
- 파일명: 모델담당자_YYYYMMDD_HHMMSS.xlsx

### 4.5 검증허브 자동화 구현

#### `POST /api/open-testhub`
Selenium WebDriver를 사용하여 검증허브 자동화:

**주요 기능**:
1. Chrome WebDriver 경로 찾기
   ```python
   def find_chrome_driver():
       # 1. 프로젝트 폴더의 drivers 디렉토리 확인
       project_driver = os.path.join(os.path.dirname(__file__), 'drivers', 'chromedriver.exe')
       if os.path.exists(project_driver):
           return project_driver

       # 2. Chrome 설치 경로에서 chromedriver.exe 찾기
       chrome_base_paths = [
           'C:\\Program Files\\Google\\Chrome\\Application',
           'C:\\Program Files (x86)\\Google\\Chrome\\Application',
       ]

       for base_path in chrome_base_paths:
           if os.path.exists(base_path):
               pattern = os.path.join(base_path, '*', 'chromedriver.exe')
               drivers = glob.glob(pattern)
               if drivers:
                   return drivers[0]

       return None
   ```

2. 다운로드 경로 설정 (실제 환경에 맞게 수정 필요)
   ```python
   # Windows 기본 다운로드 폴더
   download_dir = r'C:\Users\woonjoung.lee\Downloads'
   # 프로젝트 폴더 경로
   project_dir = r'C:\Users\woonjoung.lee\PycharmProjects\FP_Manager'
   ```

3. Chrome 옵션 설정
   ```python
   chrome_options = ChromeOptions()
   # 팝업 및 알림 비활성화
   chrome_options.add_argument('--disable-popup-blocking')
   chrome_options.add_argument('--disable-notifications')
   chrome_options.add_argument('--disable-blink-features=AutomationControlled')
   chrome_options.add_argument('--disable-dev-shm-usage')
   chrome_options.add_argument('--no-sandbox')

   # 다운로드 폴더 설정 (자동 다운로드, 팝업 없음)
   prefs = {
       'download.default_directory': download_dir,
       'download.prompt_for_download': False,
       'download.directory_upgrade': True,
       'safebrowsing.enabled': False,
       'profile.default_content_settings.popups': 0,
       'profile.default_content_setting_values.automatic_downloads': 1,
       'profile.content_settings.exceptions.automatic_downloads.*.setting': 1
   }
   chrome_options.add_experimental_option('prefs', prefs)
   chrome_options.add_experimental_option('excludeSwitches', ['enable-automation', 'enable-logging'])
   chrome_options.add_experimental_option('useAutomationExtension', False)
   ```

4. 백그라운드 스레드에서 자동화 실행:
   ```python
   def automate_clicks():
       global last_testhub_result

       try:
           # 1. 다운로드 폴더와 프로젝트 폴더의 기존 TGVerifyDetailList 파일 삭제
           folders_to_clean = [download_dir, project_dir]
           for folder in folders_to_clean:
               if os.path.exists(folder):
                   for file in os.listdir(folder):
                       if file.startswith('TGVerifyDetailList') and file.endswith(('.xlsx', '.xls')):
                           file_path = os.path.join(folder, file)
                           os.remove(file_path)

           # 2. CDP 명령으로 다운로드 동작 설정
           driver.execute_cdp_cmd('Page.setDownloadBehavior', {
               'behavior': 'allow',
               'downloadPath': download_dir
           })

           # 3. URL 열기
           driver.get(url)
           time.sleep(3)  # Chrome 로딩 대기

           # 4. XPath로 요소 순차 클릭
           wait = WebDriverWait(driver, 20)

           initial_xpaths = [
               '//*[@id="openSpan"]/img',
               '//*[@id="multi_select_ulBody_searchStdTestItemId"]/li[5]/a/input',
               '//*[@id="multi_select_ulBody_searchStdTestItemId"]/li[6]/a/input',
               '//*[@id="multi_select_btnOk_searchStdTestItemId"]',
               '//*[@id="searchBtn"]'  # 검색 버튼
           ]

           for xpath in initial_xpaths:
               element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
               element.click()
               time.sleep(1)  # 각 클릭 후 대기

           # 5. 검색 결과 로딩 대기 및 엑셀 버튼 클릭
           excel_btn_xpath = '//*[@id="excelBtn"]/span'
           excel_btn = wait.until(EC.element_to_be_clickable((By.XPATH, excel_btn_xpath)))
           time.sleep(5)  # 엑셀 버튼 클릭 전 5초 대기
           excel_btn.click()
   ```

5. 다운로드 완료 대기 (최대 60초)
   ```python
   def wait_for_download_complete(download_dir, timeout=60):
       time.sleep(3)  # 다운로드 시작 대기

       start_time = time.time()
       last_file_size = -1
       stable_count = 0

       while time.time() - start_time < timeout:
           target_files = []
           for file in os.listdir(download_dir):
               if file.startswith('TGVerifyDetailList') and file.endswith(('.xlsx', '.xls')):
                   if not file.endswith('.crdownload') and not file.endswith('.tmp'):
                       file_path = os.path.join(download_dir, file)
                       target_files.append(file_path)

           if target_files:
               file_path = target_files[0]
               current_size = os.path.getsize(file_path)
               if current_size == last_file_size:
                   stable_count += 1
                   if stable_count >= 3:  # 3번 연속 크기가 같으면 완료
                       return True
               else:
                   stable_count = 0
               last_file_size = current_size

           time.sleep(1)

       return False
   ```

6. 다운로드된 파일 처리
   ```python
   # 파일 찾기
   excel_files = []
   for file in os.listdir(download_dir):
       if (file.startswith('TGVerifyDetailList') and
           file.endswith(('.xlsx', '.xls')) and
           not file.startswith('~$')):
           file_path = os.path.join(download_dir, file)
           excel_files.append((file_path, os.path.getmtime(file_path)))

   # 가장 최근 파일 선택
   downloaded_file = max(excel_files, key=lambda x: x[1])[0]

   # 프로젝트 폴더로 파일 복사
   os.makedirs(project_dir, exist_ok=True)
   filename = os.path.basename(downloaded_file)
   project_file = os.path.join(project_dir, filename)
   shutil.copy2(downloaded_file, project_file)

   # ExcelProcessor로 처리
   processor = ExcelProcessor()
   week_number = processor.get_week_number()
   output_filename = f"{week_number}_FA_과제일정.xlsx"
   output_path = os.path.join(OUTPUT_FOLDER, output_filename)

   success = processor.process(project_file, MODEL_DB_PATH, output_path)

   if success and processor.schedule_df is not None:
       # 결과를 전역 변수에 저장
       last_testhub_result = {
           'success': True,
           'data': {
               'columns': processor.schedule_df.columns.tolist(),
               'rows': processor.schedule_df.fillna('').astype(str).values.tolist(),
               'total_rows': len(processor.schedule_df)
           },
           'filename': output_filename,
           'download_url': f'/download/{output_filename}',
           'timestamp': time.time()
       }
   ```

7. 브라우저는 닫지 않고 계속 열어둠 (사용자가 계속 사용할 수 있도록)
   ```python
   # driver.quit()  # 주석 처리하여 브라우저를 닫지 않음
   ```

8. 별도 스레드에서 실행
   ```python
   thread = threading.Thread(target=automate_clicks)
   thread.daemon = True
   thread.start()
   ```

#### `GET /api/testhub-result`
- 검증허브 처리 결과 조회
- last_testhub_result 반환

#### `POST /api/open-output-folder`
- 현재 주차의 결과 파일 다운로드
- 파일명: W{주차}_FA_과제일정.xlsx
- 파일이 없으면 404 에러

---

## 5. templates/index.html 구현

### 5.1 HTML 구조

```html
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>Monday</title>
    <style>
        /* CSS 스타일 */
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📅 Monday</h1>
            <p>Field Protocol 일정 관리 및 모델 데이터베이스</p>
        </div>

        <div class="tabs">
            <button class="tab active" onclick="switchTab('schedule')">📋 Schedule</button>
            <button class="tab" onclick="switchTab('modelInfo')">🔍 Model Info.</button>
        </div>

        <!-- Schedule 탭 -->
        <div id="scheduleTab" class="tab-content active">
            <!-- 내용 -->
        </div>

        <!-- Model Info 탭 -->
        <div id="modelInfoTab" class="tab-content">
            <!-- 내용 -->
        </div>
    </div>

    <script>
        /* JavaScript 코드 */
    </script>
</body>
</html>
```

### 5.2 CSS 스타일링

**디자인 컨셉**:
- 그라데이션 배경: `linear-gradient(135deg, #667eea 0%, #764ba2 100%)`
- 화이트 컨테이너: 둥근 모서리 (border-radius: 20px), 그림자 효과
- 탭 스타일: 활성 탭은 하단에 3px 파란색 라인
- 테이블: sticky header, 가운데 정렬, hover 효과
- 버튼: 그라데이션, hover 시 약간 위로 이동 및 그림자 효과

**주요 스타일 클래스**:
- `.container`: 메인 컨테이너
- `.header`: 헤더 (그라데이션 배경)
- `.tabs`: 탭 버튼 영역
- `.tab`: 개별 탭 버튼
- `.tab-content`: 탭 컨텐츠
- `.section-title`: 섹션 제목
- `.form-group`: 폼 그룹
- `.btn`: 기본 버튼
- `.btn-primary`, `.btn-success`, `.btn-info`, `.btn-delete`: 버튼 변형
- `.table-container`: 테이블 컨테이너 (스크롤 가능)
- `thead`, `tbody`: 테이블 스타일
- `.message`: 메시지 박스 (success/error)
- `.loading`: 로딩 스피너

**Model Info 테이블 컬럼 너비 (비율 1.4:1.2:1:1:1)**:
- 과제명: 168px
- 개발모델명: 144px
- 검증PL: 120px
- 모델담당자: 120px
- AP/CP: 120px

### 5.3 Schedule 탭 구현

**레이아웃**:
- 2열 그리드 레이아웃 (grid-template-columns: 1fr 1fr)

**왼쪽: 검증허브 의뢰조회**
- 제목: "🔍 검증허브 의뢰조회"
- 버튼: "검증허브 열기" (onclick: openTestHub())

**오른쪽: 일정 파일 업로드**
- 제목: "📊 일정 파일 업로드"
- 파일 입력: schedule_file (accept: .xlsx, .xls)
- 제출 버튼: "📥 일정 처리 시작"

**결과 섹션**:
- 로딩 스피너
- 메시지 박스
- 결과 테이블 (미리보기)
- "📁 결과 파일" 버튼 (onclick: openOutputFolder())

### 5.4 Model Info 탭 구현

**검색 섹션**:
- 검색 입력창 (Enter 키 지원)
- "검색" 버튼 (onclick: searchModel())
- "전체 보기" 버튼 (onclick: loadAllModels())

**Export 버튼**:
- "📊 Excel로 내보내기" (onclick: exportModelDB())

**테이블**:
- 헤더: 과제명, 개발모델명, 검증PL, 모델담당자, AP/CP, 작업
- contenteditable="true" (편집 가능)
- onblur: updateCell(this) (수정 시 자동 저장)
- 삭제 버튼: 🗑️ (onclick: deleteRow(rowIndex))

**추가 폼**:
- 제목: "➕ 새 모델 정보 추가"
- 5개 입력 필드 (grid layout)
- "✅추가하기" 버튼 (onclick: addModelInfo())

### 5.5 JavaScript 함수 구현

#### 전역 변수
```javascript
let currentScheduleDownloadUrl = '';
let currentSchedulePreviewUrl = '';
let schedulePreviewData = null;
```

#### Schedule 탭 함수

**`async openTestHub()`**:
1. `/api/open-testhub` POST 요청
2. 성공 시 메시지 표시
3. 2초마다 `/api/testhub-result` 조회 (최대 20회)
4. 결과가 준비되면 displaySchedulePreview() 호출

**`switchTab(tabName)`**:
- 탭 전환
- Model Info 탭 열 때 loadAllModels() 자동 호출

**파일 업로드 폼 제출**:
1. FormData 생성
2. `/upload` POST 요청
3. 성공 시 loadSchedulePreview() 호출

**`async loadSchedulePreview()`**:
- currentSchedulePreviewUrl로 미리보기 데이터 가져오기

**`showSchedulePreview()`**:
- schedulePreviewData를 테이블로 렌더링

**`displaySchedulePreview(data)`**:
- 검증허브 결과 표시

**`async openOutputFolder()`**:
- `/api/open-output-folder` POST 요청
- 결과 파일 다운로드 URL 열기

#### Model Info 탭 함수

**`async loadAllModels()`**:
- `/api/model-info/all` GET 요청
- displayModelData() 호출

**`async searchModel()`**:
- 검색어 가져오기
- `/api/model-info/search` POST 요청
- displayModelData() 호출

**`displayModelData(data)`**:
- 테이블에 데이터 렌더링
- contenteditable 셀 생성
- 삭제 버튼 추가

**`async updateCell(cell)`**:
- 셀 편집 후 자동 저장
- `/api/model-info/update` POST 요청

**`async deleteRow(rowIndex)`**:
- 확인 창 표시
- `/api/model-info/delete` POST 요청
- 성공 시 loadAllModels() 호출

**`async addModelInfo()`**:
- 입력 필드 값 검증
- `/api/model-info/add` POST 요청
- 성공 시 폼 초기화 및 loadAllModels() 호출

**`async exportModelDB()`**:
- `/api/model-info/export` 새 창으로 열기

**`showMessage(elementId, type, text)`**:
- 메시지 박스 표시
- 5초 후 자동 숨김

---

## 6. excel_scraper_merger.py 구현 (선택적)

웹 스크래핑 및 Excel 병합 기능:

### 클래스 구조: `ExcelScraperMerger`

**속성**:
- `urls`: 다운로드할 파일의 URL 딕셔너리
- `download_dir`: 다운로드 디렉토리
- `output_dir`: 출력 디렉토리

**메서드**:
- `download_excel_file(url, filename)`: URL에서 Excel 파일 다운로드
- `process_excel_file(filepath, sheet_name)`: Excel 파일 읽기 및 정리
- `add_source_column(df, source_name)`: 출처 컬럼 추가
- `download_all_files()`: 모든 URL에서 파일 다운로드
- `process_all_files(downloaded_files)`: 모든 파일 처리
- `merge_dataframes(dataframes)`: 여러 데이터프레임 병합
- `save_to_excel(df, filename)`: Excel 파일로 저장
- `run(merge)`: 전체 프로세스 실행

---

## 7. 실행 방법

### 7.1 서버 시작

```bash
python app.py
```

서버는 `http://localhost:5000` (또는 `0.0.0.0:5000`)에서 실행됩니다.

### 7.2 시작 메시지

```
============================================================
Monday 시작
접속 주소: http://localhost:5000
============================================================
```

---

## 8. 주요 특징 및 기능

### 8.1 Excel 처리
- **자동 헤더 감지**: 첫 10개 행에서 헤더 위치 자동 찾기
- **컬럼 별칭 지원**: 다양한 컬럼명 변형 대응
- **Fallback 엔진**: 여러 엔진으로 파일 읽기 시도
- **자동 정렬**: 검증항목 우선순위 → PRA 오름차순
- **VLOOKUP**: 개발모델명 기준 모델담당자 매칭
- **특별 규칙**: 송수화 시험은 모델담당자 '이운정' 자동 설정

### 8.2 모델 데이터베이스
- **CRUD 기능**: 조회, 추가, 수정, 삭제
- **검색 기능**: 과제명 또는 개발모델명으로 검색
- **인라인 편집**: contenteditable로 직접 수정
- **Excel Export**: 모델담당자 데이터베이스 내보내기

### 8.3 검증허브 자동화
- **Selenium 자동화**: Chrome WebDriver로 웹 자동화
- **자동 다운로드**: 파일 자동 다운로드 및 처리
- **다운로드 대기**: 파일 크기가 안정화될 때까지 대기
- **자동 처리**: 다운로드 완료 후 자동으로 Excel 처리
- **결과 polling**: 처리 결과를 주기적으로 조회

### 8.4 웹 인터페이스
- **반응형 디자인**: 모던한 그라데이션 디자인
- **탭 UI**: Schedule과 Model Info 분리
- **실시간 미리보기**: 처리 결과 실시간 표시
- **로딩 인디케이터**: 처리 중 스피너 표시
- **메시지 시스템**: 성공/실패 메시지 자동 표시

---

## 9. 에러 처리

### 9.1 파일 처리 에러
- 파일이 없거나 형식이 잘못된 경우
- 필요한 컬럼이 누락된 경우 (빈 값으로 처리)
- 엑셀 엔진 오류 (fallback 엔진으로 재시도)

### 9.2 데이터베이스 에러
- 데이터베이스 파일이 없는 경우 (자동 생성)
- 컬럼 순서가 잘못된 경우 (자동 재정렬)

### 9.3 검증허브 자동화 에러
- WebDriver를 찾을 수 없는 경우 (자동 다운로드 시도)
- 요소 클릭 실패 (계속 진행)
- 다운로드 타임아웃 (오류 메시지 반환)

### 9.4 웹 API 에러
- 400: 잘못된 요청 (필수 파라미터 누락)
- 403: 권한 없음 (경로 탈출 시도)
- 404: 파일을 찾을 수 없음
- 500: 서버 내부 오류

---

## 10. 보안 고려사항

### 10.1 파일 업로드
- 파일 크기 제한: 16MB
- 허용된 확장자만 업로드: .xlsx, .xls
- 안전한 파일명 생성: werkzeug.secure_filename 사용 (일부 경로)

### 10.2 경로 검증
- 다운로드 시 경로 탈출 방지
- 절대 경로로 변환 후 검증
- OUTPUT_FOLDER 외부 접근 차단

### 10.3 입력 검증
- 모든 API 요청 파라미터 검증
- 빈 값 및 타입 확인
- XSS 방지 (HTML escape)

---

## 11. 로깅

### 11.1 로그 레벨
- INFO: 정상 동작 (파일 읽기, 처리, 저장)
- WARNING: 경고 (누락된 컬럼, 매칭 실패)
- ERROR: 오류 (파일 읽기 실패, 처리 실패)

### 11.2 로그 형식
```python
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)
```

### 11.3 주요 로그 메시지
- "파일 업로드 요청 수신"
- "파일 저장됨: {filepath}"
- "처리 완료: {output_path}"
- "모델담당자 매칭 완료: {matched_count}/{total_count} 행"
- "검증허브 다운로드 파일 처리 완료"

---

## 12. 테스트 및 디버깅

### 12.1 수동 테스트
1. 서버 시작: `python app.py`
2. 브라우저로 `http://localhost:5000` 접속
3. Schedule 탭에서 파일 업로드 테스트
4. Model Info 탭에서 CRUD 테스트
5. 검증허브 자동화 테스트 (WebDriver 필요)

### 12.2 디버깅
- Flask debug 모드: `app.run(debug=True)`
- 브라우저 개발자 도구: 네트워크 탭에서 API 요청/응답 확인
- 서버 로그: 콘솔 출력 확인
- Excel 파일 직접 열어서 확인

---

## 13. 배포 및 운영

### 13.1 배포 전 체크리스트
- [ ] SECRET_KEY를 환경 변수로 설정
- [ ] TESTHUB_URL을 실제 URL로 변경
- [ ] debug=False로 설정
- [ ] 프로덕션 WSGI 서버 사용 (gunicorn, uwsgi 등)
- [ ] HTTPS 설정
- [ ] 파일 업로드 크기 제한 조정
- [ ] 로그 레벨 및 로그 파일 설정

### 13.2 프로덕션 실행 예시
```bash
gunicorn -w 4 -b 0.0.0.0:5000 app:app
```

### 13.3 백업
- db/model_manager.xlsx: 정기적으로 백업
- output 폴더: 필요한 파일만 보관

---

## 14. 확장 가능성

### 14.1 추가 기능 아이디어
- 사용자 인증 및 권한 관리
- 파일 업로드 히스토리
- 데이터 통계 및 차트
- 이메일 알림
- 스케줄러 (주기적 작업)
- API 문서 (Swagger)

### 14.2 성능 개선
- 대용량 파일 처리 최적화
- 캐싱 (Redis)
- 비동기 처리 (Celery)
- 데이터베이스 전환 (SQLite, PostgreSQL)

---

## 15. 추가 정보

### 15.1 사용된 주요 라이브러리
- **Flask**: 웹 프레임워크
- **Pandas**: 데이터 처리
- **OpenPyXL**: Excel 파일 쓰기 (xlsx)
- **xlrd**: Excel 파일 읽기 (xls)
- **Selenium**: 웹 브라우저 자동화
- **webdriver-manager**: Chrome WebDriver 자동 관리

### 15.2 참고 자료
- Flask 공식 문서: https://flask.palletsprojects.com/
- Pandas 공식 문서: https://pandas.pydata.org/
- Selenium 공식 문서: https://www.selenium.dev/

---

## 16. 구현 체크리스트

### Phase 1: 기본 구조
- [ ] 프로젝트 폴더 구조 생성
- [ ] requirements.txt 작성
- [ ] 폴더 자동 생성 로직

### Phase 2: Excel 처리
- [ ] ExcelProcessor 클래스 구현
- [ ] 파일 읽기 및 헤더 감지
- [ ] 컬럼 추출 및 별칭 처리
- [ ] 정렬 로직
- [ ] VLOOKUP 로직
- [ ] 파일 저장

### Phase 3: Flask 서버
- [ ] Flask 앱 설정
- [ ] 라우트 구현 (/, /upload, /download, /preview)
- [ ] Model Info API 구현 (CRUD)
- [ ] 검증허브 자동화 구현
- [ ] 에러 처리

### Phase 4: 웹 UI
- [ ] HTML 구조
- [ ] CSS 스타일링
- [ ] JavaScript 함수 구현
- [ ] Schedule 탭
- [ ] Model Info 탭
- [ ] 메시지 시스템

### Phase 5: 테스트 및 디버깅
- [ ] 파일 업로드 테스트
- [ ] Excel 처리 테스트
- [ ] Model Info CRUD 테스트
- [ ] 검증허브 자동화 테스트
- [ ] 에러 처리 테스트

### Phase 6: 문서화 및 배포
- [ ] README.md 작성
- [ ] 주석 추가
- [ ] 배포 설정
- [ ] 운영 가이드

---

## 17. 특별 구현 사항

### 17.1 송수화 시험 특별 처리
```python
# excel_processor.py의 vlookup_model_manager 메서드에서
if '검증항목' in schedule_df.columns:
    songsuha_mask = schedule_df['검증항목'].astype(str).str.contains('송수화', na=False)
    if songsuha_mask.sum() > 0:
        schedule_df.loc[songsuha_mask, '모델담당자'] = '이운정'
```

### 17.2 주차 계산
```python
# ISO 8601 주차 기준
now = datetime.now()
week_number = now.isocalendar()[1]
return f"W{week_number:02d}"
```

### 17.3 다운로드 완료 대기
```python
# 파일 크기가 3번 연속 동일하면 완료로 판단
if current_size == last_file_size:
    stable_count += 1
    if stable_count >= 3:
        return True
```

### 17.4 경로 검증
```python
# 디렉토리 탈출 방지
abs_output_folder = os.path.abspath(OUTPUT_FOLDER)
abs_filepath = os.path.abspath(filepath)
if not abs_filepath.startswith(abs_output_folder):
    return error
```

---

## 18. 환경 설정 및 주의사항

### 18.1 Windows 환경 특화 사항

**다운로드 경로 설정** (app.py 내 `open_testhub` 함수):
```python
# 실제 사용자 환경에 맞게 수정 필요
download_dir = r'C:\Users\{YOUR_USERNAME}\Downloads'
project_dir = r'C:\Users\{YOUR_USERNAME}\PycharmProjects\{PROJECT_NAME}'
```

**Chrome WebDriver 경로**:
- 자동 감지: `drivers/chromedriver.exe` 또는 Chrome 설치 경로
- 수동 설치: https://chromedriver.chromium.org/downloads 에서 Chrome 버전에 맞는 드라이버 다운로드
- 설치 위치: 프로젝트의 `drivers` 폴더

### 18.2 검증허브 URL 설정

**환경 변수 또는 코드에서 설정**:
```python
# app.py
TESTHUB_URL = os.getenv('TESTHUB_URL', 'https://your-actual-testhub-url.com')
```

실제 사용 시 검증허브의 정확한 URL로 변경 필요

### 18.3 필수 import 문

**app.py 상단**:
```python
#!/usr/bin/env python3
"""
Flask 웹 서버 - Monday
Schedule 관리 및 Model Info 데이터베이스 관리
"""

import os
import pandas as pd
from flask import Flask, render_template, request, send_file, jsonify, send_from_directory
from werkzeug.utils import secure_filename
from excel_processor import ExcelProcessor
import logging
from datetime import datetime
import subprocess
import platform
import webbrowser
import time
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
```

### 18.4 실행 시 출력 메시지

**app.py 마지막 부분**:
```python
if __name__ == '__main__':
    print("=" * 60)
    print("Monday 시작")
    print("접속 주소: http://localhost:5000")
    print("=" * 60)
    app.run(host='0.0.0.0', port=5000, debug=True)
```

### 18.5 데이터베이스 초기 상태

**db/model_manager.xlsx**:
- 파일이 없으면 자동 생성됨
- 초기 컬럼: ['과제명', '개발모델명', '검증PL', '모델담당자', 'AP/CP']
- 빈 데이터프레임으로 시작

### 18.6 포트 및 호스트 설정

**기본 설정**:
- Host: `0.0.0.0` (모든 네트워크 인터페이스에서 접속 가능)
- Port: `5000`
- 접속 URL: `http://localhost:5000` 또는 `http://127.0.0.1:5000`

**방화벽 설정**:
- Windows 방화벽에서 포트 5000 허용 필요 (외부 접속 시)

### 18.7 파일 인코딩

**Excel 파일**:
- 한글 파일명 지원
- UTF-8 인코딩 사용
- 셀 내용도 한글 지원

**HTML 파일**:
```html
<meta charset="UTF-8">
```

### 18.8 로깅 설정

**기본 로깅**:
```python
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)
```

**로그 출력 위치**:
- 콘솔 (표준 출력)
- 필요 시 파일 로깅 추가 가능

### 18.9 에러 메시지 처리

**사용자 친화적 에러 메시지**:
- 한글 메시지 사용
- 구체적인 문제 상황 설명
- 해결 방법 제시

예시:
```python
return jsonify({
    'success': False,
    'message': 'Chrome WebDriver를 찾을 수 없습니다.\n\n다음 중 하나를 수행해주세요:\n1. https://chromedriver.chromium.org/downloads 에서 Chrome 버전에 맞는 WebDriver를 다운로드하여 프로젝트의 drivers 폴더에 chromedriver.exe로 저장\n2. 인터넷에 연결하여 자동 다운로드 허용'
}), 500
```

### 18.10 시작 전 체크리스트

구현 시작 전 확인사항:
- [ ] Python 3.8 이상 설치 확인
- [ ] pip 최신 버전 업데이트
- [ ] Chrome 브라우저 설치 (Selenium 자동화용)
- [ ] 프로젝트 폴더 생성
- [ ] 가상환경 설정 (선택사항)
  ```bash
  python -m venv venv
  venv\Scripts\activate  # Windows
  source venv/bin/activate  # Linux/Mac
  ```
- [ ] requirements.txt 의존성 설치
  ```bash
  pip install -r requirements.txt
  ```

### 18.11 개발 vs 프로덕션 설정

**개발 환경**:
```python
app.run(host='127.0.0.1', port=5000, debug=True)
```

**프로덕션 환경**:
```python
app.run(host='0.0.0.0', port=5000, debug=False)
```

또는 gunicorn 사용:
```bash
gunicorn -w 4 -b 0.0.0.0:5000 --timeout 120 app:app
```

### 18.12 타임아웃 설정

**주요 타임아웃 값**:
- 파일 다운로드 대기: 60초
- Selenium WebDriverWait: 20초
- Chrome 로딩 대기: 3초
- 엑셀 버튼 클릭 전 대기: 5초
- 각 요소 클릭 후 대기: 1초
- 결과 polling 간격: 2초 (최대 20회 = 40초)

### 18.13 메모리 관리

**대용량 파일 처리 시**:
- 파일 크기 제한: 16MB (Flask MAX_CONTENT_LENGTH)
- 필요 시 제한 조정 가능
- 파일 처리 후 메모리 정리

**브라우저 관리**:
- Selenium 브라우저는 의도적으로 닫지 않음 (사용자가 계속 사용 가능)
- 필요 시 수동으로 종료

---

## 19. 완전한 코드 플로우

### Schedule 파일 업로드 전체 흐름

1. **사용자 액션**: 웹 UI에서 Schedule 파일 선택 및 "일정 처리 시작" 클릭
2. **프론트엔드**: FormData 생성 후 `/upload` POST 요청
3. **백엔드 (app.py)**:
   - 파일 유효성 검사 (`allowed_file`)
   - 파일 저장 (`uploads/schedule_{filename}`)
   - ExcelProcessor 인스턴스 생성
   - 주차 계산 (`get_week_number()`)
   - 처리 실행 (`processor.process(schedule_file, MODEL_DB_PATH, output_path)`)
4. **excel_processor.py**:
   - Schedule 파일 읽기 (`read_schedule_file`)
     - 헤더 자동 감지 (`find_header_row`)
     - 필요한 컬럼 추출 (부분 일치 및 별칭 지원)
   - 모델담당자 DB 읽기 (`read_model_manager_file`)
   - 데이터 정렬 (`sort_schedule_data`)
     - 검증항목 우선순위 적용
     - PRA 오름차순 정렬
   - VLOOKUP 매칭 (`vlookup_model_manager`)
     - 개발모델명으로 매칭
     - 송수화 시험 특별 처리 ('이운정' 설정)
   - 결과 저장 (`save_result`)
     - openpyxl로 저장
     - 모든 셀 가운데 정렬 적용
5. **백엔드 (app.py)**: 성공 응답 반환
   ```json
   {
     "success": true,
     "message": "파일 처리가 완료되었습니다.",
     "filename": "W47_FA_과제일정.xlsx",
     "download_url": "/download/W47_FA_과제일정.xlsx",
     "preview_url": "/preview/W47_FA_과제일정.xlsx"
   }
   ```
6. **프론트엔드**:
   - 성공 메시지 표시
   - 미리보기 데이터 로드 (`loadSchedulePreview()`)
   - 결과 테이블 표시 (`showSchedulePreview()`)

### 검증허브 자동화 전체 흐름

1. **사용자 액션**: "검증허브 열기" 버튼 클릭
2. **프론트엔드**: `/api/open-testhub` POST 요청
3. **백엔드 (app.py)**:
   - Chrome WebDriver 경로 찾기 (`find_chrome_driver()`)
   - Chrome 옵션 설정 (다운로드 폴더, 팝업 비활성화 등)
   - 별도 스레드에서 자동화 실행 (`automate_clicks()`)
     - 기존 TGVerifyDetailList 파일 삭제
     - 검증허브 URL 열기
     - XPath로 요소 순차 클릭 (5개 요소)
     - 엑셀 다운로드 버튼 클릭
     - 다운로드 완료 대기 (`wait_for_download_complete()`)
     - 파일 복사 및 처리
     - 결과를 `last_testhub_result`에 저장
   - 즉시 성공 응답 반환 (백그라운드 처리)
4. **프론트엔드**:
   - 성공 메시지 표시
   - 2초마다 `/api/testhub-result` polling (최대 20회)
5. **백엔드 (app.py)**: polling 요청에 대해 `last_testhub_result` 반환
6. **프론트엔드**:
   - 결과가 준비되면 미리보기 표시 (`displaySchedulePreview()`)

---

## 20. 최종 구현 순서 권장

구현 시 다음 순서를 권장합니다:

1. **기본 환경 설정** (1-2시간)
   - 프로젝트 폴더 생성
   - requirements.txt 작성 및 패키지 설치
   - 폴더 구조 생성

2. **excel_processor.py 구현** (3-4시간)
   - 클래스 및 기본 메서드
   - 파일 읽기 로직
   - 정렬 및 VLOOKUP 로직
   - 파일 저장 로직
   - 테스트 (단독 실행)

3. **app.py 기본 구현** (2-3시간)
   - Flask 앱 설정
   - 기본 라우트 (/, /upload, /download, /preview)
   - 유틸리티 함수
   - 테스트 (파일 업로드)

4. **templates/index.html 기본 UI** (2-3시간)
   - HTML 구조
   - CSS 스타일링
   - Schedule 탭 기본 기능
   - 테스트 (파일 업로드 및 미리보기)

5. **Model Info API 구현** (2-3시간)
   - 데이터베이스 CRUD 함수
   - API 라우트 구현
   - 프론트엔드 Model Info 탭
   - 테스트 (CRUD 작업)

6. **검증허브 자동화 구현** (3-4시간)
   - Selenium 설정
   - Chrome WebDriver 관리
   - 자동화 로직
   - 다운로드 및 처리 로직
   - 테스트 (전체 플로우)

7. **통합 테스트 및 디버깅** (2-3시간)
   - 전체 기능 테스트
   - 에러 케이스 처리
   - 로깅 확인
   - 사용자 경험 개선

8. **문서화 및 배포** (1-2시간)
   - README.md 작성
   - 주석 추가
   - 배포 준비

**총 예상 시간**: 16-24시간 (중급 개발자 기준)

---

이 프롬프트를 따라 구현하면 현재 Monday 웹서버와 동일한 기능을 가진 시스템을 구축할 수 있습니다.
