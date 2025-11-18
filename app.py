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


if __name__ == '__main__':
    print("=" * 60)
    print("Monday 시작")
    print("접속 주소: http://localhost:5000")
    print("=" * 60)
    app.run(host='0.0.0.0', port=5000, debug=True)
