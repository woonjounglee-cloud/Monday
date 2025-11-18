# Edge WebDriver 설치 안내

검증허브 자동 클릭 기능을 사용하려면 Edge WebDriver가 필요합니다.

## 설치 방법

1. 현재 설치된 Edge 버전 확인
   - Edge 브라우저 열기 → 설정 → 정보 → 버전 확인
   - 예: 142.0.3595.x

2. WebDriver 다운로드
   - https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/
   - Edge 버전과 **동일한 버전**의 WebDriver 다운로드
   - 예: Edge 142.0.3595 → WebDriver 142.0.3595

3. 설치
   - 다운로드한 zip 파일 압축 해제
   - `msedgedriver.exe` 파일을 이 폴더(`drivers/`)에 복사
   - 파일명이 정확히 `msedgedriver.exe`인지 확인

## 파일 구조

```
Monday/
  ├── drivers/
  │   ├── msedgedriver.exe  ← 여기에 파일 복사
  │   └── README.md
  ├── app.py
  └── ...
```

## 확인 방법

파일을 올바르게 복사했다면 "검증허브 의뢰조회" 버튼이 정상 작동합니다.
