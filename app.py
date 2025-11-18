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

# 로깅 설정
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Flask 앱 초기화
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

# 검증허브 URL (사용자가 수정 가능)
TESTHUB_URL = os.getenv('TESTHUB_URL', 'https://example.com')  # 실제 URL로 변경 필요

# 검증허브 처리 결과 저장용 전역 변수
last_testhub_result = None


def allowed_file(filename):
    """허용된 파일 확장자인지 확인"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def load_model_database():
    """모델담당자 데이터베이스 로드"""
    try:
        # 올바른 컬럼 순서 정의
        correct_columns = ['과제명', '개발모델명', '검증PL', '모델담당자', 'AP/CP']

        if os.path.exists(MODEL_DB_PATH):
            df = pd.read_excel(MODEL_DB_PATH, engine='openpyxl')

            # 컬럼이 모두 존재하는지 확인
            missing_cols = [col for col in correct_columns if col not in df.columns]
            if missing_cols:
                logger.warning(f"누락된 컬럼: {missing_cols}")
                for col in missing_cols:
                    df[col] = ''

            # 컬럼 순서를 올바르게 재정렬
            df = df[correct_columns]
            return df
        else:
            # 데이터베이스 파일이 없으면 빈 데이터프레임 생성
            df = pd.DataFrame(columns=correct_columns)
            df.to_excel(MODEL_DB_PATH, index=False, engine='openpyxl')
            logger.info(f"새 모델담당자 데이터베이스 생성: {MODEL_DB_PATH}")
            return df
    except Exception as e:
        logger.error(f"모델담당자 데이터베이스 로드 실패: {e}")
        return pd.DataFrame(columns=['과제명', '개발모델명', '검증PL', '모델담당자', 'AP/CP'])


def save_model_database(df):
    """모델담당자 데이터베이스 저장"""
    try:
        # 올바른 컬럼 순서 정의
        correct_columns = ['과제명', '개발모델명', '검증PL', '모델담당자', 'AP/CP']

        # 컬럼 순서를 올바르게 재정렬
        if all(col in df.columns for col in correct_columns):
            df = df[correct_columns]

        df.to_excel(MODEL_DB_PATH, index=False, engine='openpyxl')
        logger.info(f"모델담당자 데이터베이스 저장 완료: {len(df)} 행")
        return True
    except Exception as e:
        logger.error(f"모델담당자 데이터베이스 저장 실패: {e}")
        return False


@app.route('/')
def index():
    """메인 페이지"""
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_files():
    """
    파일 업로드 및 처리 엔드포인트 - 단일 Schedule 파일 처리
    """
    try:
        logger.info("파일 업로드 요청 수신")

        # Schedule 파일 확인
        if 'schedule_file' not in request.files:
            return jsonify({
                'success': False,
                'message': '일정 파일을 업로드해주세요.'
            }), 400

        file = request.files['schedule_file']

        if not file or not file.filename:
            return jsonify({
                'success': False,
                'message': '일정 파일을 업로드해주세요.'
            }), 400

        if not allowed_file(file.filename):
            return jsonify({
                'success': False,
                'message': '엑셀 파일(.xlsx, .xls)만 업로드 가능합니다.'
            }), 400

        # 안전한 파일명 생성
        filename = secure_filename(file.filename)
        filepath = os.path.join(UPLOAD_FOLDER, f"schedule_{filename}")

        # 파일 저장
        file.save(filepath)
        logger.info(f"파일 저장됨: {filepath}")

        # ExcelProcessor로 처리
        processor = ExcelProcessor()

        # 주차 계산
        week_number = processor.get_week_number()
        output_filename = f"{week_number}_FA_과제일정.xlsx"
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)

        # 처리 실행 (단일 파일, db/model_manager.xlsx 사용)
        success = processor.process(filepath, MODEL_DB_PATH, output_path)

        if success:
            logger.info(f"처리 완료: {output_path}")

            # 업로드된 파일 정리 (선택사항)
            # if os.path.exists(filepath):
            #     os.remove(filepath)

            return jsonify({
                'success': True,
                'message': '파일 처리가 완료되었습니다.',
                'filename': output_filename,
                'download_url': f'/download/{output_filename}',
                'preview_url': f'/preview/{output_filename}'
            })
        else:
            return jsonify({
                'success': False,
                'message': '파일 처리 중 오류가 발생했습니다.'
            }), 500

    except Exception as e:
        logger.error(f"업로드 처리 중 오류: {e}")
        import traceback
        logger.error(traceback.format_exc())
        return jsonify({
            'success': False,
            'message': f'오류 발생: {str(e)}'
        }), 500


@app.route('/preview/<filename>')
def preview_file(filename):
    """
    결과 파일 미리보기 엔드포인트
    """
    try:
        filepath = os.path.join(OUTPUT_FOLDER, filename)

        if not os.path.exists(filepath):
            return jsonify({
                'success': False,
                'message': '파일을 찾을 수 없습니다.'
            }), 404

        # 엑셀 파일 읽기
        df = pd.read_excel(filepath, engine='openpyxl')

        # 데이터프레임을 딕셔너리로 변환
        data = {
            'columns': df.columns.tolist(),
            'rows': df.fillna('').astype(str).values.tolist(),
            'total_rows': len(df)
        }

        return jsonify({
            'success': True,
            'data': data
        })

    except Exception as e:
        logger.error(f"미리보기 중 오류: {e}")
        import traceback
        logger.error(traceback.format_exc())
        return jsonify({
            'success': False,
            'message': f'오류 발생: {str(e)}'
        }), 500


@app.route('/download/<path:filename>')
def download_file(filename):
    """
    결과 파일 다운로드 엔드포인트
    """
    try:
        logger.info(f"다운로드 요청: {filename}")

        # 파일 경로 생성 (secure_filename 사용하지 않음 - 한글 파일명 지원)
        filepath = os.path.join(OUTPUT_FOLDER, filename)

        # 절대 경로로 변환
        abs_output_folder = os.path.abspath(OUTPUT_FOLDER)
        abs_filepath = os.path.abspath(filepath)

        # 경로 검증 (보안: 디렉토리 탈출 방지)
        if not abs_filepath.startswith(abs_output_folder):
            logger.error(f"잘못된 경로 접근 시도: {filepath}")
            return jsonify({
                'success': False,
                'message': '잘못된 경로입니다.'
            }), 403

        if not os.path.exists(abs_filepath):
            logger.error(f"파일을 찾을 수 없음: {abs_filepath}")
            logger.info(f"OUTPUT_FOLDER 내용: {os.listdir(OUTPUT_FOLDER)}")
            return jsonify({
                'success': False,
                'message': '파일을 찾을 수 없습니다.'
            }), 404

        logger.info(f"파일 다운로드 시작: {abs_filepath}")

        # send_file 사용 (절대 경로)
        return send_file(
            abs_filepath,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        logger.error(f"다운로드 중 오류: {e}")
        import traceback
        logger.error(traceback.format_exc())
        return jsonify({
            'success': False,
            'message': f'오류 발생: {str(e)}'
        }), 500


# ============= Model Info API =============

@app.route('/api/model-info/all', methods=['GET'])
def get_all_model_info():
    """모든 모델 정보 가져오기"""
    try:
        df = load_model_database()
        data = {
            'success': True,
            'data': {
                'columns': df.columns.tolist(),
                'rows': df.fillna('').astype(str).values.tolist(),
                'total_rows': len(df)
            }
        }
        return jsonify(data)
    except Exception as e:
        logger.error(f"모델 정보 조회 실패: {e}")
        return jsonify({
            'success': False,
            'message': f'오류 발생: {str(e)}'
        }), 500


@app.route('/api/model-info/search', methods=['POST'])
def search_model_info():
    """모델 정보 검색"""
    try:
        data = request.get_json()
        keyword = data.get('keyword', '').strip()

        if not keyword:
            return jsonify({
                'success': False,
                'message': '검색어를 입력해주세요.'
            }), 400

        df = load_model_database()

        # 과제명 또는 개발모델명에서 검색
        mask = (df['과제명'].astype(str).str.contains(keyword, case=False, na=False) |
                df['개발모델명'].astype(str).str.contains(keyword, case=False, na=False))

        result_df = df[mask]

        return jsonify({
            'success': True,
            'data': {
                'columns': result_df.columns.tolist(),
                'rows': result_df.fillna('').astype(str).values.tolist(),
                'total_rows': len(result_df)
            },
            'message': f'{len(result_df)}개의 결과를 찾았습니다.'
        })

    except Exception as e:
        logger.error(f"모델 정보 검색 실패: {e}")
        import traceback
        logger.error(traceback.format_exc())
        return jsonify({
            'success': False,
            'message': f'오류 발생: {str(e)}'
        }), 500


@app.route('/api/model-info/add', methods=['POST'])
def add_model_info():
    """모델 정보 추가"""
    try:
        data = request.get_json()

        # 필수 필드 확인
        required_fields = ['과제명', '개발모델명', '검증PL', '모델담당자', 'AP/CP']
        for field in required_fields:
            if field not in data or not data[field].strip():
                return jsonify({
                    'success': False,
                    'message': f'{field}을(를) 입력해주세요.'
                }), 400

        # 데이터베이스 로드
        df = load_model_database()

        # 새 행 추가
        new_row = {
            '과제명': data['과제명'].strip(),
            '개발모델명': data['개발모델명'].strip(),
            '검증PL': data['검증PL'].strip(),
            '모델담당자': data['모델담당자'].strip(),
            'AP/CP': data['AP/CP'].strip()
        }

        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

        # 저장
        if save_model_database(df):
            return jsonify({
                'success': True,
                'message': '모델 정보가 추가되었습니다.',
                'data': new_row
            })
        else:
            return jsonify({
                'success': False,
                'message': '저장 중 오류가 발생했습니다.'
            }), 500

    except Exception as e:
        logger.error(f"모델 정보 추가 실패: {e}")
        import traceback
        logger.error(traceback.format_exc())
        return jsonify({
            'success': False,
            'message': f'오류 발생: {str(e)}'
        }), 500


@app.route('/api/model-info/export', methods=['GET'])
def export_model_info():
    """모델담당자 데이터베이스를 Excel 파일로 내보내기"""
    try:
        if not os.path.exists(MODEL_DB_PATH):
            return jsonify({
                'success': False,
                'message': '데이터베이스 파일을 찾을 수 없습니다.'
            }), 404

        # 현재 시간으로 파일명 생성
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        export_filename = f'모델담당자_{timestamp}.xlsx'
        export_path = os.path.join(OUTPUT_FOLDER, export_filename)

        # 데이터베이스 복사
        df = load_model_database()
        df.to_excel(export_path, index=False, engine='openpyxl')

        logger.info(f"모델담당자 데이터 export: {export_path}")

        # 파일 전송
        return send_file(
            export_path,
            as_attachment=True,
            download_name=export_filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        logger.error(f"Export 중 오류: {e}")
        import traceback
        logger.error(traceback.format_exc())
        return jsonify({
            'success': False,
            'message': f'오류 발생: {str(e)}'
        }), 500


@app.route('/api/model-info/update', methods=['POST'])
def update_model_info():
    """모델 정보 수정"""
    try:
        data = request.get_json()

        # 필수 필드 확인
        required_fields = ['row_index', 'field', 'value']
        for field in required_fields:
            if field not in data:
                return jsonify({
                    'success': False,
                    'message': f'{field}이(가) 없습니다.'
                }), 400

        row_index = int(data['row_index'])
        field_name = data['field']
        new_value = data['value'].strip()

        # 데이터베이스 로드
        df = load_model_database()

        # 인덱스 범위 확인
        if row_index < 0 or row_index >= len(df):
            return jsonify({
                'success': False,
                'message': '유효하지 않은 행 번호입니다.'
            }), 400

        # 필드명 확인
        if field_name not in df.columns:
            return jsonify({
                'success': False,
                'message': '유효하지 않은 필드명입니다.'
            }), 400

        # 값 업데이트
        df.at[row_index, field_name] = new_value

        # 저장
        if save_model_database(df):
            return jsonify({
                'success': True,
                'message': '모델 정보가 수정되었습니다.',
                'data': {
                    'row_index': row_index,
                    'field': field_name,
                    'value': new_value
                }
            })
        else:
            return jsonify({
                'success': False,
                'message': '저장 중 오류가 발생했습니다.'
            }), 500

    except Exception as e:
        logger.error(f"모델 정보 수정 실패: {e}")
        import traceback
        logger.error(traceback.format_exc())
        return jsonify({
            'success': False,
            'message': f'오류 발생: {str(e)}'
        }), 500


@app.route('/api/model-info/delete', methods=['POST'])
def delete_model_info():
    """모델 정보 삭제"""
    try:
        data = request.get_json()

        # 필수 필드 확인
        if 'row_index' not in data:
            return jsonify({
                'success': False,
                'message': 'row_index가 없습니다.'
            }), 400

        row_index = int(data['row_index'])

        # 데이터베이스 로드
        df = load_model_database()

        # 인덱스 범위 확인
        if row_index < 0 or row_index >= len(df):
            return jsonify({
                'success': False,
                'message': '유효하지 않은 행 번호입니다.'
            }), 400

        # 행 삭제
        df = df.drop(index=row_index).reset_index(drop=True)

        # 저장
        if save_model_database(df):
            return jsonify({
                'success': True,
                'message': '모델 정보가 삭제되었습니다.'
            })
        else:
            return jsonify({
                'success': False,
                'message': '저장 중 오류가 발생했습니다.'
            }), 500

    except Exception as e:
        logger.error(f"모델 정보 삭제 실패: {e}")
        import traceback
        logger.error(traceback.format_exc())
        return jsonify({
            'success': False,
            'message': f'오류 발생: {str(e)}'
        }), 500


@app.route('/api/open-testhub', methods=['POST'])
def open_testhub():
    """검증허브 URL을 Chrome으로 열고 자동으로 요소 클릭"""
    logger.info("=== 검증허브 열기 요청 수신 ===")
    try:
        url = TESTHUB_URL
        logger.info(f"TestHub URL: {url}")

        # Chrome WebDriver 경로 찾기
        def find_chrome_driver():
            """Chrome WebDriver 경로를 찾는 함수"""
            import glob

            # 1. 프로젝트 폴더의 drivers 디렉토리 확인
            project_driver = os.path.join(os.path.dirname(__file__), 'drivers', 'chromedriver.exe')
            if os.path.exists(project_driver):
                logger.info(f"프로젝트 폴더의 드라이버 사용: {project_driver}")
                return project_driver

            # 2. Chrome 설치 경로에서 chromedriver.exe 찾기
            chrome_base_paths = [
                'C:\\Program Files\\Google\\Chrome\\Application',
                'C:\\Program Files (x86)\\Google\\Chrome\\Application',
            ]

            for base_path in chrome_base_paths:
                if os.path.exists(base_path):
                    # 버전 폴더들을 검색
                    pattern = os.path.join(base_path, '*', 'chromedriver.exe')
                    drivers = glob.glob(pattern)
                    if drivers:
                        driver_path = drivers[0]  # 첫 번째 발견된 드라이버 사용
                        logger.info(f"Chrome 설치 폴더의 드라이버 사용: {driver_path}")
                        return driver_path

            return None

        # 다운로드 경로 설정 (Windows 기본 다운로드 폴더)
        download_dir = r'C:\Users\woonjoung.lee\Downloads'
        # 프로젝트 폴더 경로
        project_dir = r'C:\Users\woonjoung.lee\PycharmProjects\FP_Manager'

        # Chrome WebDriver 설정
        chrome_options = ChromeOptions()
        # 브라우저를 백그라운드에서 실행하지 않음 (사용자가 볼 수 있도록)
        # chrome_options.add_argument('--headless')  # 주석 처리하여 화면에 표시

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
            'safebrowsing.enabled': False,  # 안전 확인 비활성화
            'profile.default_content_settings.popups': 0,
            'profile.default_content_setting_values.automatic_downloads': 1,
            'profile.content_settings.exceptions.automatic_downloads.*.setting': 1
        }
        chrome_options.add_experimental_option('prefs', prefs)
        chrome_options.add_experimental_option('excludeSwitches', ['enable-automation', 'enable-logging'])
        chrome_options.add_experimental_option('useAutomationExtension', False)

        # WebDriver 초기화
        driver = None
        driver_path = find_chrome_driver()
        logger.info(f"Chrome WebDriver 검색 결과: {driver_path}")

        try:
            if driver_path:
                # 찾은 드라이버 경로 사용
                service = ChromeService(executable_path=driver_path)
                driver = webdriver.Chrome(service=service, options=chrome_options)
                logger.info(f"WebDriver 초기화 성공: {driver_path}")
            else:
                # 드라이버를 찾지 못한 경우 webdriver-manager 사용
                try:
                    service = ChromeService(ChromeDriverManager().install())
                    driver = webdriver.Chrome(service=service, options=chrome_options)
                    logger.info("webdriver-manager로 Chrome WebDriver 초기화 성공")
                except Exception as wdm_error:
                    logger.error(f"webdriver-manager 실패: {wdm_error}")
                    return jsonify({
                        'success': False,
                        'message': 'Chrome WebDriver를 찾을 수 없습니다.\n\n다음 중 하나를 수행해주세요:\n1. https://chromedriver.chromium.org/downloads 에서 Chrome 버전에 맞는 WebDriver를 다운로드하여 프로젝트의 drivers 폴더에 chromedriver.exe로 저장\n2. 인터넷에 연결하여 자동 다운로드 허용'
                    }), 500

        except Exception as e:
            logger.error(f"WebDriver 초기화 실패: {e}")
            return jsonify({
                'success': False,
                'message': f'Chrome WebDriver를 초기화할 수 없습니다.\n\n오류: {str(e)}\n\nhttps://chromedriver.chromium.org/downloads 에서 Chrome 버전에 맞는 WebDriver를 다운로드하여 프로젝트의 drivers 폴더에 chromedriver.exe로 저장해주세요.'
            }), 500

        # 백그라운드에서 실행 (비동기)
        import threading

        def automate_clicks():
            global last_testhub_result  # 함수 시작 부분에 global 선언
            try:
                # 1. 다운로드 폴더와 프로젝트 폴더의 기존 TGVerifyDetailList 파일 삭제
                folders_to_clean = [download_dir, project_dir]
                for folder in folders_to_clean:
                    try:
                        if os.path.exists(folder):
                            deleted_count = 0
                            for file in os.listdir(folder):
                                if file.startswith('TGVerifyDetailList') and file.endswith(('.xlsx', '.xls')):
                                    file_path = os.path.join(folder, file)
                                    os.remove(file_path)
                                    logger.info(f"기존 파일 삭제: {file_path}")
                                    deleted_count += 1
                            if deleted_count > 0:
                                logger.info(f"{folder}: {deleted_count}개 파일 삭제 완료")
                            else:
                                logger.info(f"{folder}: 삭제할 파일 없음")
                    except Exception as del_error:
                        logger.warning(f"{folder} 파일 삭제 실패 (계속 진행): {del_error}")

                # CDP 명령으로 다운로드 동작 설정 (자동 다운로드, 팝업 없음)
                try:
                    driver.execute_cdp_cmd('Page.setDownloadBehavior', {
                        'behavior': 'allow',
                        'downloadPath': download_dir
                    })
                    logger.info(f"다운로드 동작 설정 완료: {download_dir}")
                except Exception as cdp_error:
                    logger.warning(f"CDP 명령 실패 (계속 진행): {cdp_error}")

                # URL 열기
                driver.get(url)
                logger.info(f"TestHub URL 열기: {url}")

                # Chrome 열린 후 3초 대기
                logger.info("Chrome 로딩 대기 중 (3초)...")
                time.sleep(3)
                logger.info("대기 완료, XPath 클릭 시작")

                # 페이지 로드 대기
                wait = WebDriverWait(driver, 20)

                # XPath 목록 (순서대로 클릭)
                # 처음 5개는 순차적으로 클릭, 마지막은 별도 처리
                initial_xpaths = [
                    '//*[@id="openSpan"]/img',
                    '//*[@id="multi_select_ulBody_searchStdTestItemId"]/li[5]/a/input',
                    '//*[@id="multi_select_ulBody_searchStdTestItemId"]/li[6]/a/input',
                    '//*[@id="multi_select_btnOk_searchStdTestItemId"]',
                    '//*[@id="searchBtn"]'  # 검색 버튼
                ]

                # 처음 5개 요소를 순서대로 클릭
                for i, xpath in enumerate(initial_xpaths, 1):
                    try:
                        # 요소가 클릭 가능할 때까지 대기
                        element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
                        element.click()
                        logger.info(f"클릭 완료 ({i}/6): {xpath}")

                        # 각 클릭 후 대기 (페이지 반응 대기)
                        time.sleep(1)

                    except Exception as e:
                        logger.error(f"요소 클릭 실패 ({i}/6): {xpath} - {e}")
                        # 클릭 실패해도 계속 진행

                # searchBtn 클릭 후 화면이 바뀔 때까지 대기
                # excelBtn이 클릭 가능해질 때까지 최대 20초 대기
                logger.info("검색 버튼 클릭 완료. 검색 결과 로딩 대기 중...")
                try:
                    excel_btn_xpath = '//*[@id="excelBtn"]/span'
                    excel_btn = wait.until(EC.element_to_be_clickable((By.XPATH, excel_btn_xpath)))
                    logger.info("검색 결과 로드 완료. 엑셀 버튼 클릭 가능")

                    # 엑셀 버튼 클릭 전 5초 추가 대기
                    logger.info("엑셀 버튼 클릭 전 5초 대기 중...")
                    time.sleep(5)

                    # 엑셀 다운로드 버튼 클릭
                    excel_btn.click()
                    logger.info(f"클릭 완료 (6/6): {excel_btn_xpath}")
                    logger.info("엑셀 다운로드 시작됨")

                except Exception as e:
                    logger.error(f"엑셀 버튼 클릭 실패: {e}")

                logger.info("자동 클릭 완료 (엑셀 다운로드 버튼 포함)")

                # 다운로드 완료 대기 함수
                def wait_for_download_complete(download_dir, timeout=60):
                    """다운로드가 완료될 때까지 대기"""
                    logger.info("다운로드 완료 대기 중...")

                    # 다운로드 시작 대기
                    time.sleep(3)

                    start_time = time.time()
                    last_file_size = -1
                    stable_count = 0

                    while time.time() - start_time < timeout:
                        if not os.path.exists(download_dir):
                            time.sleep(0.5)
                            continue

                        # TGVerifyDetailList로 시작하는 파일 찾기
                        target_files = []
                        for file in os.listdir(download_dir):
                            if file.startswith('TGVerifyDetailList') and file.endswith(('.xlsx', '.xls')):
                                if not file.endswith('.crdownload') and not file.endswith('.tmp'):
                                    file_path = os.path.join(download_dir, file)
                                    target_files.append(file_path)

                        if target_files:
                            # 파일이 발견되면 크기가 안정화될 때까지 대기
                            file_path = target_files[0]
                            try:
                                current_size = os.path.getsize(file_path)
                                if current_size == last_file_size:
                                    stable_count += 1
                                    if stable_count >= 3:  # 3번 연속 크기가 같으면 완료
                                        logger.info(f"다운로드 완료: {file_path} ({current_size} bytes)")
                                        return True
                                else:
                                    stable_count = 0
                                    logger.info(f"다운로드 진행 중... ({current_size} bytes)")
                                last_file_size = current_size
                            except Exception as e:
                                logger.warning(f"파일 크기 확인 실패: {e}")

                        time.sleep(1)

                    logger.warning("다운로드 타임아웃")
                    return False

                # 다운로드 완료 대기 (최대 60초)
                download_complete = wait_for_download_complete(download_dir, timeout=60)

                if not download_complete:
                    logger.error("다운로드가 완료되지 않았습니다")
                    last_testhub_result = {
                        'success': False,
                        'message': '파일 다운로드 시간이 초과되었습니다.',
                        'timestamp': time.time()
                    }
                    return

                # 다운로드된 파일 찾기 (TGVerifyDetailList로 시작하는 파일)
                download_path = download_dir
                if os.path.exists(download_path):
                    # TGVerifyDetailList로 시작하는 가장 최근 엑셀 파일 찾기
                    excel_files = []
                    for file in os.listdir(download_path):
                        if (file.startswith('TGVerifyDetailList') and
                            file.endswith(('.xlsx', '.xls')) and
                            not file.startswith('~$')):
                            file_path = os.path.join(download_path, file)
                            excel_files.append((file_path, os.path.getmtime(file_path)))

                    logger.info(f"TGVerifyDetailList 파일 검색 결과: {len(excel_files)}개 발견")

                    if excel_files:
                        # 가장 최근 파일 선택
                        downloaded_file = max(excel_files, key=lambda x: x[1])[0]
                        logger.info(f"다운로드된 파일 발견: {downloaded_file}")

                        # 프로젝트 폴더로 파일 복사
                        try:
                            os.makedirs(project_dir, exist_ok=True)
                            filename = os.path.basename(downloaded_file)
                            project_file = os.path.join(project_dir, filename)

                            shutil.copy2(downloaded_file, project_file)
                            logger.info(f"파일을 프로젝트 폴더로 복사: {downloaded_file} → {project_file}")

                            # 복사된 파일로 처리
                            latest_file = project_file
                        except Exception as copy_error:
                            logger.error(f"파일 복사 실패: {copy_error}")
                            # 복사 실패 시 다운로드 경로의 파일 사용
                            latest_file = downloaded_file

                        # 파일 처리
                        try:
                            from excel_processor import ExcelProcessor

                            processor = ExcelProcessor()

                            # 출력 파일명 생성 (주차 계산)
                            week_number = processor.get_week_number()
                            output_filename = f"{week_number}_FA_과제일정.xlsx"
                            output_path = os.path.join(OUTPUT_FOLDER, output_filename)
                            logger.info(f"출력 파일 경로: {output_path}")

                            # 처리 실행 (process 메서드 사용)
                            success = processor.process(latest_file, MODEL_DB_PATH, output_path)

                            if success and processor.schedule_df is not None:
                                logger.info("검증허브 다운로드 파일 처리 완료")

                                # 결과를 전역 변수에 저장 (클라이언트가 조회할 수 있도록)
                                # 데이터프레임을 columns, rows 형식으로 변환 (/preview와 동일한 형식)
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
                            else:
                                logger.error("파일 처리 실패")
                                last_testhub_result = {
                                    'success': False,
                                    'message': '파일 처리 중 오류가 발생했습니다.',
                                    'timestamp': time.time()
                                }
                        except Exception as proc_error:
                            logger.error(f"파일 처리 중 오류: {proc_error}")
                            import traceback
                            logger.error(traceback.format_exc())
                            last_testhub_result = {
                                'success': False,
                                'message': f'파일 처리 중 오류: {str(proc_error)}',
                                'timestamp': time.time()
                            }
                    else:
                        logger.warning("다운로드된 엑셀 파일을 찾을 수 없습니다")
                        last_testhub_result = {
                            'success': False,
                            'message': '다운로드된 엑셀 파일을 찾을 수 없습니다.',
                            'timestamp': time.time()
                        }
                else:
                    logger.warning(f"다운로드 폴더가 존재하지 않습니다: {download_path}")
                    last_testhub_result = {
                        'success': False,
                        'message': f'다운로드 폴더가 존재하지 않습니다: {download_path}',
                        'timestamp': time.time()
                    }

                # 브라우저는 열어둠 (사용자가 계속 사용할 수 있도록)
                # driver.quit()  # 주석 처리하여 브라우저를 닫지 않음

            except Exception as e:
                logger.error(f"자동화 실행 중 오류: {e}")
                import traceback
                logger.error(traceback.format_exc())
                # driver.quit()

        # 별도 스레드에서 자동화 실행
        thread = threading.Thread(target=automate_clicks)
        thread.daemon = True
        thread.start()

        return jsonify({
            'success': True,
            'message': f'검증허브를 열고 자동 클릭을 시작했습니다.',
            'url': url
        })

    except Exception as e:
        logger.error(f"TestHub 열기 실패: {e}")
        import traceback
        logger.error(traceback.format_exc())
        return jsonify({
            'success': False,
            'message': f'오류 발생: {str(e)}'
        }), 500


@app.route('/api/testhub-result', methods=['GET'])
def get_testhub_result():
    """검증허브 처리 결과 조회"""
    global last_testhub_result

    if last_testhub_result is None:
        return jsonify({
            'ready': False,
            'message': '처리 중이거나 결과가 없습니다.'
        })

    result = last_testhub_result.copy()
    result['ready'] = True

    # 결과 반환 후 초기화 (한 번만 조회 가능)
    # last_testhub_result = None

    return jsonify(result)


@app.route('/api/open-output-folder', methods=['POST'])
def open_output_folder():
    """결과 파일이 저장된 output 폴더를 탐색기로 열기"""
    try:
        # OUTPUT_FOLDER의 절대 경로 구하기
        abs_output_folder = os.path.abspath(OUTPUT_FOLDER)

        # 폴더가 존재하는지 확인
        if not os.path.exists(abs_output_folder):
            return jsonify({
                'success': False,
                'message': f'폴더를 찾을 수 없습니다: {abs_output_folder}'
            }), 404

        # 운영체제에 따라 폴더 열기
        system = platform.system()

        if system == 'Windows':
            # Windows: explorer를 사용하여 폴더 열기
            os.startfile(abs_output_folder)
            logger.info(f"Windows 탐색기로 폴더 열기: {abs_output_folder}")
        elif system == 'Darwin':  # macOS
            subprocess.Popen(['open', abs_output_folder])
            logger.info(f"macOS Finder로 폴더 열기: {abs_output_folder}")
        else:  # Linux
            subprocess.Popen(['xdg-open', abs_output_folder])
            logger.info(f"Linux 파일 관리자로 폴더 열기: {abs_output_folder}")

        return jsonify({
            'success': True,
            'message': f'폴더를 열었습니다: {abs_output_folder}',
            'folder_path': abs_output_folder
        })

    except Exception as e:
        logger.error(f"폴더 열기 실패: {e}")
        import traceback
        logger.error(traceback.format_exc())
        return jsonify({
            'success': False,
            'message': f'폴더 열기 실패: {str(e)}'
        }), 500


if __name__ == '__main__':
    print("=" * 60)
    print("Monday 시작")
    print("접속 주소: http://localhost:5000")
    print("=" * 60)
    app.run(host='0.0.0.0', port=5000, debug=True)
