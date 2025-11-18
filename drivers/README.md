# Chrome WebDriver 설치 안내

검증허브 자동 클릭 기능을 사용하려면 Chrome WebDriver가 필요합니다.

## 설치 방법

1. 현재 설치된 Chrome 버전 확인
   - Chrome 브라우저 열기 → 설정 → Chrome 정보 → 버전 확인
   - 또는 주소창에 `chrome://version/` 입력
   - 예: 131.0.6778.x

2. WebDriver 다운로드
   - https://googlechromelabs.github.io/chrome-for-testing/
   - 또는 https://chromedriver.chromium.org/downloads
   - Chrome 버전과 **동일한 버전**의 ChromeDriver 다운로드
   - 예: Chrome 131.0.6778 → ChromeDriver 131.0.6778

3. 설치
   - 다운로드한 zip 파일 압축 해제
   - `chromedriver.exe` 파일을 이 폴더(`drivers/`)에 복사
   - 파일명이 정확히 `chromedriver.exe`인지 확인

## 파일 구조

```
Monday/
  ├── drivers/
  │   ├── chromedriver.exe  ← 여기에 파일 복사
  │   └── README.md
  ├── app.py
  └── ...
```

## 자동 설치 (인터넷 연결 필요)

drivers 폴더에 chromedriver.exe가 없으면 webdriver-manager가 자동으로 다운로드를 시도합니다.
인터넷 연결이 불가능한 경우 위의 수동 설치 방법을 사용하세요.

## 확인 방법

파일을 올바르게 복사했다면 "검증허브 의뢰조회" 버튼이 정상 작동합니다.
