#!/usr/bin/env python3
"""
Excel Scraper and Merger
특정 웹사이트에서 엑셀 파일을 다운로드하고 데이터를 가공하여 하나의 파일로 병합하는 프로그램
"""

import os
import requests
import pandas as pd
from datetime import datetime
from typing import List, Dict
import logging

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class ExcelScraperMerger:
    """엑셀 파일 다운로드 및 병합 클래스"""

    def __init__(self):
        """초기화"""
        # 다운로드할 파일의 URL 설정
        self.urls = {
            'Field_Protocol_Driving_Test': '',  # 여기에 URL을 입력하세요
            'Field_Protocol_Stationary_Call_Test': '',  # 여기에 URL을 입력하세요
            'Field_Protocol_VQ_Test': ''  # 여기에 URL을 입력하세요
        }

        # 다운로드 디렉토리 설정
        self.download_dir = 'downloads'
        self.output_dir = 'output'

        # 디렉토리 생성
        os.makedirs(self.download_dir, exist_ok=True)
        os.makedirs(self.output_dir, exist_ok=True)

    def download_excel_file(self, url: str, filename: str) -> str:
        """
        URL에서 엑셀 파일을 다운로드

        Args:
            url: 다운로드할 파일의 URL
            filename: 저장할 파일명

        Returns:
            str: 다운로드된 파일의 경로
        """
        if not url:
            logger.warning(f"URL이 설정되지 않았습니다: {filename}")
            return None

        try:
            logger.info(f"다운로드 시작: {filename}")

            # 파일 다운로드
            response = requests.get(url, timeout=30)
            response.raise_for_status()

            # 파일 저장
            filepath = os.path.join(self.download_dir, filename)
            with open(filepath, 'wb') as f:
                f.write(response.content)

            logger.info(f"다운로드 완료: {filepath}")
            return filepath

        except requests.exceptions.RequestException as e:
            logger.error(f"다운로드 실패 ({filename}): {e}")
            return None

    def process_excel_file(self, filepath: str, sheet_name: str = None) -> pd.DataFrame:
        """
        엑셀 파일을 읽어서 데이터프레임으로 반환

        Args:
            filepath: 엑셀 파일 경로
            sheet_name: 읽을 시트 이름 (None이면 첫 번째 시트)

        Returns:
            pd.DataFrame: 처리된 데이터프레임
        """
        if not filepath or not os.path.exists(filepath):
            logger.warning(f"파일이 존재하지 않습니다: {filepath}")
            return None

        try:
            logger.info(f"파일 처리 중: {filepath}")

            # 엑셀 파일 읽기
            df = pd.read_excel(filepath, sheet_name=sheet_name)

            # 기본 데이터 정리
            # 빈 행 제거
            df = df.dropna(how='all')

            # 빈 열 제거
            df = df.dropna(axis=1, how='all')

            logger.info(f"처리 완료: {len(df)} 행, {len(df.columns)} 열")
            return df

        except Exception as e:
            logger.error(f"파일 처리 실패 ({filepath}): {e}")
            return None

    def add_source_column(self, df: pd.DataFrame, source_name: str) -> pd.DataFrame:
        """
        데이터프레임에 출처(source) 컬럼 추가

        Args:
            df: 원본 데이터프레임
            source_name: 출처 이름

        Returns:
            pd.DataFrame: 출처 컬럼이 추가된 데이터프레임
        """
        if df is not None:
            df_copy = df.copy()
            df_copy.insert(0, 'Source', source_name)
            return df_copy
        return None

    def download_all_files(self) -> Dict[str, str]:
        """
        모든 URL에서 파일 다운로드

        Returns:
            Dict[str, str]: 파일명과 파일 경로의 딕셔너리
        """
        downloaded_files = {}

        for name, url in self.urls.items():
            filename = f"{name}.xlsx"
            filepath = self.download_excel_file(url, filename)
            if filepath:
                downloaded_files[name] = filepath

        return downloaded_files

    def process_all_files(self, downloaded_files: Dict[str, str]) -> List[pd.DataFrame]:
        """
        다운로드된 모든 파일 처리

        Args:
            downloaded_files: 파일명과 경로의 딕셔너리

        Returns:
            List[pd.DataFrame]: 처리된 데이터프레임 리스트
        """
        processed_data = []

        for name, filepath in downloaded_files.items():
            df = self.process_excel_file(filepath)
            if df is not None:
                # 출처 컬럼 추가
                df = self.add_source_column(df, name)
                processed_data.append(df)

        return processed_data

    def merge_dataframes(self, dataframes: List[pd.DataFrame]) -> pd.DataFrame:
        """
        여러 데이터프레임을 하나로 병합

        Args:
            dataframes: 병합할 데이터프레임 리스트

        Returns:
            pd.DataFrame: 병합된 데이터프레임
        """
        if not dataframes:
            logger.warning("병합할 데이터프레임이 없습니다")
            return None

        try:
            logger.info(f"{len(dataframes)}개의 데이터프레임 병합 중...")

            # 모든 데이터프레임을 세로로 결합
            merged_df = pd.concat(dataframes, ignore_index=True)

            logger.info(f"병합 완료: {len(merged_df)} 행")
            return merged_df

        except Exception as e:
            logger.error(f"병합 실패: {e}")
            return None

    def save_to_excel(self, df: pd.DataFrame, filename: str = None) -> str:
        """
        데이터프레임을 엑셀 파일로 저장

        Args:
            df: 저장할 데이터프레임
            filename: 파일명 (None이면 타임스탬프 사용)

        Returns:
            str: 저장된 파일 경로
        """
        if df is None:
            logger.warning("저장할 데이터가 없습니다")
            return None

        try:
            if filename is None:
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                filename = f'merged_data_{timestamp}.xlsx'

            filepath = os.path.join(self.output_dir, filename)

            logger.info(f"파일 저장 중: {filepath}")
            df.to_excel(filepath, index=False, engine='openpyxl')

            logger.info(f"저장 완료: {filepath}")
            return filepath

        except Exception as e:
            logger.error(f"파일 저장 실패: {e}")
            return None

    def run(self, merge: bool = True) -> str:
        """
        전체 프로세스 실행

        Args:
            merge: True이면 모든 파일을 병합, False이면 개별 파일로 저장

        Returns:
            str: 결과 파일 경로
        """
        logger.info("=== Excel Scraper and Merger 시작 ===")

        # 1. 모든 파일 다운로드
        logger.info("\n1. 파일 다운로드 중...")
        downloaded_files = self.download_all_files()

        if not downloaded_files:
            logger.error("다운로드된 파일이 없습니다")
            return None

        # 2. 모든 파일 처리
        logger.info("\n2. 파일 처리 중...")
        processed_data = self.process_all_files(downloaded_files)

        if not processed_data:
            logger.error("처리된 데이터가 없습니다")
            return None

        # 3. 병합 또는 개별 저장
        if merge:
            logger.info("\n3. 데이터 병합 중...")
            merged_df = self.merge_dataframes(processed_data)

            if merged_df is not None:
                logger.info("\n4. 병합된 파일 저장 중...")
                result_path = self.save_to_excel(merged_df, 'merged_field_protocols.xlsx')
            else:
                result_path = None
        else:
            logger.info("\n3. 개별 파일 저장 중...")
            result_paths = []
            for i, df in enumerate(processed_data):
                source_name = df['Source'].iloc[0] if 'Source' in df.columns else f'file_{i}'
                filepath = self.save_to_excel(df, f'{source_name}_processed.xlsx')
                if filepath:
                    result_paths.append(filepath)
            result_path = result_paths

        logger.info("\n=== 완료 ===")
        return result_path


def main():
    """메인 함수"""
    # ExcelScraperMerger 인스턴스 생성
    scraper = ExcelScraperMerger()

    # URL 설정 (실제 URL로 변경하세요)
    scraper.urls['Field_Protocol_Driving_Test'] = 'https://example.com/driving_test.xlsx'
    scraper.urls['Field_Protocol_Stationary_Call_Test'] = 'https://example.com/call_test.xlsx'
    scraper.urls['Field_Protocol_VQ_Test'] = 'https://example.com/vq_test.xlsx'

    # 실행
    # merge=True: 모든 파일을 하나로 병합
    # merge=False: 각 파일을 개별적으로 처리하여 저장
    result = scraper.run(merge=True)

    if result:
        print(f"\n성공! 결과 파일: {result}")
    else:
        print("\n실패: 파일 처리 중 오류가 발생했습니다")


if __name__ == '__main__':
    main()
