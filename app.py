#!/usr/bin/env python3
"""
Flask 웹 서버 - Excel 파일 업로드 및 병합
"""

import os
from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename
from excel_processor import ExcelProcessor
import logging

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
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

# 폴더 생성
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


def allowed_file(filename):
    """허용된 파일 확장자인지 확인"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    """메인 페이지"""
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_files():
    """
    파일 업로드 및 처리 엔드포인트
    """
    try:
        logger.info("파일 업로드 요청 수신")

        # 파일 확인
        files = {}
        file_keys = ['driving_test', 'stationary_call_test', 'vq_test', 'model_manager']

        uploaded_files = {}

        for key in file_keys:
            if key in request.files:
                file = request.files[key]
                if file and file.filename and allowed_file(file.filename):
                    # 안전한 파일명 생성
                    filename = secure_filename(file.filename)
                    filepath = os.path.join(UPLOAD_FOLDER, f"{key}_{filename}")

                    # 파일 저장
                    file.save(filepath)
                    uploaded_files[key] = filepath
                    logger.info(f"파일 저장됨: {key} -> {filepath}")

        # 최소 1개의 테스트 파일은 있어야 함
        test_files = {k: v for k, v in uploaded_files.items() if k != 'model_manager'}
        if not test_files:
            return jsonify({
                'success': False,
                'message': '최소 1개 이상의 테스트 파일을 업로드해주세요.'
            }), 400

        # ExcelProcessor로 처리
        processor = ExcelProcessor()

        # 주차 계산
        week_number = processor.get_week_number()
        output_filename = f"{week_number}_FA_과제일정.xlsx"
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)

        # 테스트 파일 딕셔너리 생성
        test_files_dict = {
            'Driving_Test': uploaded_files.get('driving_test'),
            'Stationary_Call_Test': uploaded_files.get('stationary_call_test'),
            'VQ_Test': uploaded_files.get('vq_test')
        }

        # 모델담당자 파일
        model_manager_file = uploaded_files.get('model_manager')

        # 처리 실행
        success = processor.process(test_files_dict, model_manager_file, output_path)

        if success:
            logger.info(f"처리 완료: {output_path}")

            # 업로드된 파일 정리 (선택사항)
            # for filepath in uploaded_files.values():
            #     if os.path.exists(filepath):
            #         os.remove(filepath)

            return jsonify({
                'success': True,
                'message': '파일 처리가 완료되었습니다.',
                'filename': output_filename,
                'download_url': f'/download/{output_filename}'
            })
        else:
            return jsonify({
                'success': False,
                'message': '파일 처리 중 오류가 발생했습니다.'
            }), 500

    except Exception as e:
        logger.error(f"업로드 처리 중 오류: {e}")
        return jsonify({
            'success': False,
            'message': f'오류 발생: {str(e)}'
        }), 500


@app.route('/download/<filename>')
def download_file(filename):
    """
    결과 파일 다운로드 엔드포인트
    """
    try:
        filepath = os.path.join(OUTPUT_FOLDER, filename)

        if not os.path.exists(filepath):
            return jsonify({
                'success': False,
                'message': '파일을 찾을 수 없습니다.'
            }), 404

        return send_file(
            filepath,
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        logger.error(f"다운로드 중 오류: {e}")
        return jsonify({
            'success': False,
            'message': f'오류 발생: {str(e)}'
        }), 500


if __name__ == '__main__':
    print("=" * 60)
    print("Excel 병합 서버 시작")
    print("접속 주소: http://localhost:5000")
    print("=" * 60)
    app.run(host='0.0.0.0', port=5000, debug=True)
