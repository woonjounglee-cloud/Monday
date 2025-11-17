#!/usr/bin/env python3
"""
Excel Processor for FA Task Schedule
3개의 테스트 엑셀 파일과 모델담당자 파일을 병합하여 결과 파일 생성
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

            # 먼저 헤더 없이 읽어서 헤더 위치 찾기
            df_temp = self.read_excel_with_fallback(filepath, header=None)
            header_row = self.find_header_row(df_temp, self.REQUIRED_COLUMNS)

            # 실제 데이터 읽기
            df = self.read_excel_with_fallback(
                filepath,
                header=header_row,
                dtype=str  # 모든 데이터를 문자열로 읽어서 타입 에러 방지
            )

            # 컬럼명 정리 (공백 제거)
            df.columns = df.columns.str.strip()

            logger.info(f"읽은 데이터: {len(df)} 행, {len(df.columns)} 열")
            logger.info(f"컬럼 목록: {df.columns.tolist()}")

            # 필요한 컬럼 찾기 (부분 일치 허용)
            column_mapping = {}
            for req_col in self.REQUIRED_COLUMNS:
                for df_col in df.columns:
                    if req_col in str(df_col) or str(df_col) in req_col:
                        column_mapping[df_col] = req_col
                        break

            if not column_mapping:
                logger.error(f"필요한 컬럼을 찾을 수 없습니다: {test_name}")
                logger.error(f"파일의 컬럼: {df.columns.tolist()}")
                logger.error(f"필요한 컬럼: {self.REQUIRED_COLUMNS}")
                return None

            # 컬럼명 변경
            df = df.rename(columns=column_mapping)

            # 누락된 컬럼 확인 및 추가
            missing_columns = [col for col in self.REQUIRED_COLUMNS if col not in df.columns]
            if missing_columns:
                logger.warning(f"누락된 컬럼 ({test_name}): {missing_columns}")
                for col in missing_columns:
                    df[col] = ''

            # 필요한 컬럼만 선택
            df = df[self.REQUIRED_COLUMNS].copy()

            # 빈 행 제거 (모든 값이 NaN 또는 빈 문자열인 행)
            df = df.replace('', pd.NA)
            df = df.dropna(how='all')

            # NaN을 빈 문자열로 변경
            df = df.fillna('')

            logger.info(f"추출 완료: {test_name} - {len(df)} 행")
            return df

        except Exception as e:
            logger.error(f"파일 읽기 실패 ({test_name}): {e}")
            import traceback
            logger.error(traceback.format_exc())
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

            # 데이터 읽기
            df = self.read_excel_with_fallback(
                filepath,
                dtype=str  # 모든 데이터를 문자열로 읽음
            )

            # 컬럼명 정리
            df.columns = df.columns.str.strip()

            logger.info(f"모델담당자 파일 컬럼: {df.columns.tolist()}")

            # 필요한 컬럼 확인 (개발모델명, 모델담당자, AP/CP)
            required_cols = ['개발모델명', '모델담당자', 'AP/CP']

            # 컬럼명 매핑 (부분 일치 허용)
            column_mapping = {}
            for req_col in required_cols:
                for df_col in df.columns:
                    col_lower = str(df_col).lower().strip()
                    if req_col in df_col or '개발모델' in df_col or 'model' in col_lower:
                        if '개발모델명' not in column_mapping.values():
                            column_mapping[df_col] = '개발모델명'
                    elif '모델담당자' in df_col or '담당자' in df_col:
                        if '모델담당자' not in column_mapping.values():
                            column_mapping[df_col] = '모델담당자'
                    elif 'ap/cp' in col_lower or 'apcp' in col_lower:
                        if 'AP/CP' not in column_mapping.values():
                            column_mapping[df_col] = 'AP/CP'

            if column_mapping:
                df = df.rename(columns=column_mapping)

            # 필수 컬럼 확인
            missing = [col for col in required_cols if col not in df.columns]
            if missing:
                logger.warning(f"모델담당자 파일에 누락된 컬럼: {missing}")

            # 빈 행 제거
            df = df.replace('', pd.NA)
            df = df.dropna(how='all')
            df = df.fillna('')

            logger.info(f"모델담당자 파일 읽기 완료: {len(df)} 행")
            self.model_manager_df = df
            return df

        except Exception as e:
            logger.error(f"모델담당자 파일 읽기 실패: {e}")
            import traceback
            logger.error(traceback.format_exc())
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
                # PRA 값이 있는 행만 정렬
                merged_df['PRA'] = merged_df['PRA'].astype(str)

                # 빈 값이 아닌 행만 정렬
                mask = merged_df['PRA'] != ''
                df_with_pra = merged_df[mask].copy()
                df_without_pra = merged_df[~mask].copy()

                # PRA가 있는 행 정렬
                df_with_pra = df_with_pra.sort_values(by='PRA', ascending=True)

                # 다시 합치기
                merged_df = pd.concat([df_with_pra, df_without_pra], ignore_index=True)
                logger.info("PRA 컬럼으로 정렬 완료")

            logger.info(f"병합 완료: {len(merged_df)} 행")
            return merged_df

        except Exception as e:
            logger.error(f"병합 실패: {e}")
            import traceback
            logger.error(traceback.format_exc())
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

                # 개발모델명 컬럼의 공백 제거 및 대소문자 통일
                merged_df['개발모델명_clean'] = merged_df['개발모델명'].str.strip()
                model_lookup['개발모델명_clean'] = model_lookup['개발모델명'].str.strip()

                # merge를 사용하여 매칭
                result_df = merged_df.merge(
                    model_lookup[['개발모델명_clean', '모델담당자', 'AP/CP']],
                    on='개발모델명_clean',
                    how='left',
                    suffixes=('', '_lookup')
                )

                # 매칭된 값으로 업데이트
                if '모델담당자_lookup' in result_df.columns:
                    merged_df['모델담당자'] = result_df['모델담당자_lookup'].fillna('')
                if 'AP/CP_lookup' in result_df.columns:
                    merged_df['AP/CP'] = result_df['AP/CP_lookup'].fillna('')

                matched_count = (merged_df['모델담당자'] != '').sum()
                logger.info(f"모델담당자 매칭 완료: {matched_count}/{len(merged_df)} 행")

            return merged_df

        except Exception as e:
            logger.error(f"모델담당자 매칭 실패: {e}")
            import traceback
            logger.error(traceback.format_exc())
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
            df_to_save = df[available_columns].copy()

            # 엑셀 파일로 저장 (openpyxl 엔진 사용)
            df_to_save.to_excel(output_path, index=False, engine='openpyxl')

            logger.info(f"저장 완료: {output_path} ({len(df_to_save)} 행, {len(df_to_save.columns)} 열)")
            return True

        except Exception as e:
            logger.error(f"파일 저장 실패: {e}")
            import traceback
            logger.error(traceback.format_exc())
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
            import traceback
            logger.error(traceback.format_exc())
            return False
