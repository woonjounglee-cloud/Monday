#!/usr/bin/env python3
"""
Excel Processor for FA Task Schedule
3개의 테스트 엑셀 파일과 모델담당자 파일을 병합하여 결과 파일 생성
"""

import pandas as pd
from datetime import datetime
from typing import Dict, Optional
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class ExcelProcessor:
    """엑셀 파일 처리 및 병합 클래스"""

    # 추출할 컬럼 목록
    REQUIRED_COLUMNS = [
        '검증항목', '과제명', '개발모델명', '검증단계',
        'PRA', '외뢰일', '완료요청일', '검증 PL'
    ]

    def __init__(self):
        """초기화"""
        self.test_dataframes = []
        self.model_manager_df = None

    def get_week_number(self) -> str:
        """
        현재 날짜의 주차를 계산하여 W{주차} 형식으로 반환

        Returns:
            str: W{주차} 형식 (예: W47)
        """
        now = datetime.now()
        week_number = now.isocalendar()[1]
        return f"W{week_number:02d}"

    def read_test_file(self, filepath: str, test_name: str) -> pd.DataFrame:
        """
        테스트 파일을 읽어서 필요한 컬럼만 추출

        Args:
            filepath: 엑셀 파일 경로
            test_name: 테스트 이름

        Returns:
            pd.DataFrame: 추출된 데이터프레임
        """
        try:
            logger.info(f"파일 읽기 중: {test_name} - {filepath}")

            # 엑셀 파일 읽기
            df = pd.read_excel(filepath)

            # 필요한 컬럼만 선택
            available_columns = [col for col in self.REQUIRED_COLUMNS if col in df.columns]

            if not available_columns:
                logger.warning(f"필요한 컬럼을 찾을 수 없습니다: {test_name}")
                return None

            # 누락된 컬럼 확인
            missing_columns = [col for col in self.REQUIRED_COLUMNS if col not in df.columns]
            if missing_columns:
                logger.warning(f"누락된 컬럼 ({test_name}): {missing_columns}")
                # 누락된 컬럼은 빈 값으로 추가
                for col in missing_columns:
                    df[col] = ''

            # 필요한 컬럼만 선택
            df = df[self.REQUIRED_COLUMNS].copy()

            # 빈 행 제거 (모든 값이 NaN인 행)
            df = df.dropna(how='all')

            logger.info(f"추출 완료: {test_name} - {len(df)} 행")
            return df

        except Exception as e:
            logger.error(f"파일 읽기 실패 ({test_name}): {e}")
            return None

    def read_model_manager_file(self, filepath: str) -> pd.DataFrame:
        """
        모델담당자 파일을 읽기

        Args:
            filepath: 엑셀 파일 경로

        Returns:
            pd.DataFrame: 모델담당자 데이터프레임
        """
        try:
            logger.info(f"모델담당자 파일 읽기 중: {filepath}")

            df = pd.read_excel(filepath)

            # 필요한 컬럼 확인 (개발모델명, 모델담당자, AP/CP)
            required_cols = ['개발모델명', '모델담당자', 'AP/CP']

            # 컬럼명 매핑 (파일마다 다를 수 있음)
            column_mapping = {}
            for col in df.columns:
                col_lower = col.lower().strip()
                if '개발모델명' in col or 'model' in col_lower:
                    column_mapping[col] = '개발모델명'
                elif '모델담당자' in col or '담당자' in col:
                    column_mapping[col] = '모델담당자'
                elif 'ap/cp' in col_lower or 'apcp' in col_lower:
                    column_mapping[col] = 'AP/CP'

            if column_mapping:
                df = df.rename(columns=column_mapping)

            # 필수 컬럼 확인
            missing = [col for col in required_cols if col not in df.columns]
            if missing:
                logger.warning(f"모델담당자 파일에 누락된 컬럼: {missing}")

            logger.info(f"모델담당자 파일 읽기 완료: {len(df)} 행")
            self.model_manager_df = df
            return df

        except Exception as e:
            logger.error(f"모델담당자 파일 읽기 실패: {e}")
            return None

    def add_test_file(self, filepath: str, test_name: str):
        """
        테스트 파일 추가

        Args:
            filepath: 파일 경로
            test_name: 테스트 이름
        """
        df = self.read_test_file(filepath, test_name)
        if df is not None:
            self.test_dataframes.append(df)

    def merge_test_files(self) -> pd.DataFrame:
        """
        모든 테스트 파일을 하나로 병합

        Returns:
            pd.DataFrame: 병합된 데이터프레임
        """
        if not self.test_dataframes:
            logger.warning("병합할 테스트 파일이 없습니다")
            return None

        try:
            logger.info(f"{len(self.test_dataframes)}개의 테스트 파일 병합 중...")

            # 모든 데이터프레임을 세로로 결합
            merged_df = pd.concat(self.test_dataframes, ignore_index=True)

            # PRA 컬럼으로 오름차순 정렬
            if 'PRA' in merged_df.columns:
                # PRA 컬럼을 문자열로 변환 후 정렬 (NaN 값 처리)
                merged_df['PRA'] = merged_df['PRA'].astype(str)
                merged_df = merged_df.sort_values(by='PRA', ascending=True)
                merged_df = merged_df.reset_index(drop=True)
                logger.info("PRA 컬럼으로 정렬 완료")

            logger.info(f"병합 완료: {len(merged_df)} 행")
            return merged_df

        except Exception as e:
            logger.error(f"병합 실패: {e}")
            return None

    def vlookup_model_manager(self, merged_df: pd.DataFrame) -> pd.DataFrame:
        """
        VLOOKUP 방식으로 모델담당자 정보 매칭

        Args:
            merged_df: 병합된 데이터프레임

        Returns:
            pd.DataFrame: 모델담당자 정보가 추가된 데이터프레임
        """
        if merged_df is None:
            logger.warning("병합된 데이터가 없습니다")
            return None

        if self.model_manager_df is None:
            logger.warning("모델담당자 파일이 없습니다")
            # 빈 컬럼 추가
            merged_df['모델담당자'] = ''
            merged_df['AP/CP'] = ''
            return merged_df

        try:
            logger.info("모델담당자 정보 매칭 중...")

            # 빈 컬럼 추가
            merged_df['모델담당자'] = ''
            merged_df['AP/CP'] = ''

            # 개발모델명을 기준으로 VLOOKUP
            if '개발모델명' in merged_df.columns and '개발모델명' in self.model_manager_df.columns:
                # pandas의 merge를 사용하여 VLOOKUP 구현
                model_lookup = self.model_manager_df[['개발모델명', '모델담당자', 'AP/CP']].copy()

                # 중복 제거 (첫 번째 항목만 유지)
                model_lookup = model_lookup.drop_duplicates(subset=['개발모델명'], keep='first')

                # 기존 데이터프레임에서 개발모델명 저장
                temp_df = merged_df.copy()

                # merge를 사용하여 매칭
                result_df = temp_df.merge(
                    model_lookup,
                    on='개발모델명',
                    how='left',
                    suffixes=('', '_lookup')
                )

                # 매칭된 값으로 업데이트
                if '모델담당자_lookup' in result_df.columns:
                    merged_df['모델담당자'] = result_df['모델담당자_lookup'].fillna('')
                if 'AP/CP_lookup' in result_df.columns:
                    merged_df['AP/CP'] = result_df['AP/CP_lookup'].fillna('')

                matched_count = merged_df['모델담당자'].notna().sum()
                logger.info(f"모델담당자 매칭 완료: {matched_count}/{len(merged_df)} 행")

            return merged_df

        except Exception as e:
            logger.error(f"모델담당자 매칭 실패: {e}")
            # 에러 발생 시 빈 컬럼이라도 추가
            if '모델담당자' not in merged_df.columns:
                merged_df['모델담당자'] = ''
            if 'AP/CP' not in merged_df.columns:
                merged_df['AP/CP'] = ''
            return merged_df

    def save_result(self, df: pd.DataFrame, output_path: str) -> bool:
        """
        결과를 엑셀 파일로 저장

        Args:
            df: 저장할 데이터프레임
            output_path: 저장 경로

        Returns:
            bool: 성공 여부
        """
        if df is None:
            logger.warning("저장할 데이터가 없습니다")
            return False

        try:
            logger.info(f"파일 저장 중: {output_path}")

            # 최종 컬럼 순서 정렬
            final_columns = self.REQUIRED_COLUMNS + ['모델담당자', 'AP/CP']

            # 컬럼이 모두 있는지 확인하고 순서대로 정렬
            available_columns = [col for col in final_columns if col in df.columns]
            df = df[available_columns]

            # 엑셀 파일로 저장
            df.to_excel(output_path, index=False, engine='openpyxl')

            logger.info(f"저장 완료: {output_path} ({len(df)} 행)")
            return True

        except Exception as e:
            logger.error(f"파일 저장 실패: {e}")
            return False

    def process(self, test_files: Dict[str, str], model_manager_file: Optional[str],
                output_path: str) -> bool:
        """
        전체 프로세스 실행

        Args:
            test_files: {'Driving_Test': filepath, 'Stationary_Call_Test': filepath, 'VQ_Test': filepath}
            model_manager_file: 모델담당자 파일 경로 (선택사항)
            output_path: 결과 파일 저장 경로

        Returns:
            bool: 성공 여부
        """
        try:
            logger.info("=== Excel 처리 시작 ===")

            # 1. 테스트 파일 읽기
            for test_name, filepath in test_files.items():
                if filepath:
                    self.add_test_file(filepath, test_name)

            if not self.test_dataframes:
                logger.error("처리할 테스트 파일이 없습니다")
                return False

            # 2. 테스트 파일 병합
            merged_df = self.merge_test_files()
            if merged_df is None:
                return False

            # 3. 모델담당자 파일 읽기
            if model_manager_file:
                self.read_model_manager_file(model_manager_file)

            # 4. VLOOKUP으로 모델담당자 정보 매칭
            result_df = self.vlookup_model_manager(merged_df)

            # 5. 결과 저장
            success = self.save_result(result_df, output_path)

            if success:
                logger.info("=== 처리 완료 ===")
            else:
                logger.error("=== 처리 실패 ===")

            return success

        except Exception as e:
            logger.error(f"처리 중 오류 발생: {e}")
            return False
