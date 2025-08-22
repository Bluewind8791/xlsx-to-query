import pandas as pd
import re
import os
import sys
from typing import List, Dict
from pathlib import Path
import traceback
import time


class ExcelQueryGenerator:
    def __init__(self):
        """Excel 데이터를 기반으로 쿼리를 생성하는 클래스"""
        # 실행 파일의 위치를 기준으로 경로 설정
        if getattr(sys, 'frozen', False):
            # PyInstaller로 패키징된 경우
            self.base_dir = Path(sys.executable).parent
        else:
            # 일반 Python 실행 시
            self.base_dir = Path(__file__).parent

        self.input_dir = self.base_dir / "input"
        self.output_dir = self.base_dir / "output"
        self.template_file = self.base_dir / "sample_query.txt"

    def run(self):
        """사용자 친화적인 메인 실행 함수"""
        print("=" * 60)
        print("          Excel to Query Generator v1.0")
        print("=" * 60)
        print()

        try:
            # 초기 설정 확인
            if not self.check_initial_setup():
                self.wait_for_exit()
                return

            # 쿼리 생성 실행
            self.generate_queries()

        except KeyboardInterrupt:
            print("\n\n사용자에 의해 중단되었습니다.")
        except Exception as e:
            print(f"\n예상치 못한 오류가 발생했습니다:")
            print(f"오류 내용: {str(e)}")
            print("\n상세 오류 정보:")
            traceback.print_exc()
        finally:
            self.wait_for_exit()

    def check_initial_setup(self) -> bool:
        """초기 설정 확인"""
        print("🔍 초기 설정을 확인하는 중...")

        # 디렉토리 설정
        self.setup_directories()

        # 템플릿 파일 확인
        if not self.template_file.exists():
            print(f"\n❌ 오류: 쿼리 템플릿 파일이 없습니다!")
            print(f"다음 위치에 sample_query.txt 파일을 생성해주세요:")
            print(f"📁 {self.template_file}")
            print("\n예시 내용:")
            print("update sample_table set use_yn = {사용여부} where sample_pk = {문서번호};")
            return False

        # Excel 파일 확인
        excel_files = self.find_excel_files()
        if not excel_files:
            print(f"\n⚠️  경고: Excel 파일이 없습니다!")
            print(f"다음 폴더에 처리할 Excel 파일(.xlsx, .xls)을 넣어주세요:")
            print(f"📁 {self.input_dir}")
            return False

        print(f"✅ 설정 확인 완료!")
        print(f"   - 템플릿 파일: ✓")
        print(f"   - Excel 파일: {len(excel_files)}개 발견")
        print()

        return True

    def generate_queries(self):
        """메인 실행 함수 - 모든 Excel 파일을 일괄 처리"""
        try:
            # 1. 쿼리 템플릿 읽기
            print("📖 쿼리 템플릿을 읽는 중...")
            query_template = self.read_query_template()

            # 2. input 디렉토리에서 Excel 파일 찾기
            excel_files = self.find_excel_files()

            print(f"🔄 {len(excel_files)}개의 Excel 파일을 처리합니다...\n")

            # 3. 각 Excel 파일을 처리
            processed_files = 0
            total_queries = 0

            for i, excel_file in enumerate(excel_files, 1):
                try:
                    print(f"[{i}/{len(excel_files)}] 📊 {excel_file.name}")
                    queries_count = self.process_single_excel(excel_file, query_template)
                    processed_files += 1
                    total_queries += queries_count
                    print(f"    ✅ 완료 ({queries_count}개 쿼리 생성)\n")
                except Exception as e:
                    print(f"    ❌ 실패: {str(e)}\n")
                    continue

            # 4. 결과 요약
            print("=" * 60)
            print("🎉 처리 완료!")
            print(f"   - 처리된 파일: {processed_files}/{len(excel_files)}개")
            print(f"   - 생성된 쿼리: {total_queries}개")
            print(f"   - 출력 폴더: {self.output_dir}")
            print("=" * 60)

        except Exception as e:
            print(f"❌ 오류가 발생했습니다: {str(e)}")
            raise

    def setup_directories(self):
        """필요한 디렉토리 확인 및 생성"""
        # input 디렉토리 확인
        if not self.input_dir.exists():
            self.input_dir.mkdir(parents=True, exist_ok=True)
            print(f"📁 {self.input_dir.name} 폴더를 생성했습니다.")

        # output 디렉토리 확인 및 생성
        if not self.output_dir.exists():
            self.output_dir.mkdir(parents=True, exist_ok=True)
            print(f"📁 {self.output_dir.name} 폴더를 생성했습니다.")

    def find_excel_files(self) -> List[Path]:
        """input 디렉토리에서 Excel 파일들을 찾아 반환"""
        excel_extensions = ['*.xlsx', '*.xls']
        excel_files = []

        for extension in excel_extensions:
            excel_files.extend(self.input_dir.glob(extension))

        # 파일명으로 정렬
        return sorted(excel_files)

    def process_single_excel(self, excel_file: Path, query_template: str) -> int:
        """단일 Excel 파일을 처리하고 생성된 쿼리 수를 반환"""

        # 1. Excel 파일 읽기
        data = self.read_excel_file(excel_file)

        # 2. 쿼리 생성
        queries = self.generate_queries_from_data(data, query_template)

        # 3. 출력 파일명 생성 (확장자를 .txt로 변경)
        output_filename = excel_file.stem + ".txt"
        output_path = self.output_dir / output_filename

        # 4. 쿼리를 텍스트 파일로 저장
        self.write_queries_to_file(queries, output_path)

        print(f"    📝 저장: {output_filename}")

        return len(queries)

    def read_query_template(self) -> str:
        """쿼리 템플릿 파일을 읽어서 반환"""
        try:
            with open(self.template_file, 'r', encoding='utf-8') as file:
                template = file.read().strip()

            if not template:
                raise ValueError(f"쿼리 템플릿 파일이 비어있습니다: {self.template_file}")

            # 플레이스홀더 확인
            placeholders = re.findall(r'\{([^}]+)\}', template)
            if placeholders:
                print(f"    🎯 플레이스홀더: {', '.join(placeholders)}")

            return template

        except Exception as e:
            raise Exception(f"쿼리 템플릿 파일 읽기 실패: {str(e)}")

    def read_excel_file(self, excel_file: Path) -> pd.DataFrame:
        """Excel 파일을 읽어서 DataFrame으로 반환"""
        try:
            # Excel 파일의 첫 번째 시트만 읽기
            df = pd.read_excel(excel_file, sheet_name=0, engine='openpyxl')

            if df.empty:
                raise ValueError(f"Excel 파일이 비어있습니다")

            # 컬럼명의 공백 제거
            df.columns = df.columns.str.strip()

            print(f"    📋 컬럼: {', '.join(df.columns[:3])}{'...' if len(df.columns) > 3 else ''}")
            print(f"    📊 데이터: {len(df)}행")

            return df

        except Exception as e:
            raise Exception(f"Excel 파일 읽기 실패: {str(e)}")

    def generate_queries_from_data(self, df: pd.DataFrame, template: str) -> List[str]:
        """데이터프레임과 템플릿을 기반으로 쿼리 생성"""
        queries = []

        # 템플릿에서 플레이스홀더 찾기
        placeholders = re.findall(r'\{([^}]+)\}', template)

        if not placeholders:
            raise ValueError("쿼리 템플릿에서 플레이스홀더 {}를 찾을 수 없습니다.")

        # 각 행에 대해 쿼리 생성
        for index, row in df.iterrows():
            query = template

            for placeholder in placeholders:
                # 컬럼명 매칭 (대소문자 무시)
                column_value = self.find_column_value(df.columns, row, placeholder)

                if column_value is not None:
                    # None이나 NaN 값 처리
                    if pd.isna(column_value):
                        formatted_value = "NULL"
                    else:
                        formatted_value = self.format_value(column_value)

                    # 플레이스홀더를 실제 값으로 치환
                    query = query.replace(f'{{{placeholder}}}', formatted_value)
                else:
                    query = query.replace(f'{{{placeholder}}}', "NULL")

            queries.append(query)

        return queries

    def find_column_value(self, columns: pd.Index, row: pd.Series, target_column: str):
        """컬럼명으로 해당 행의 값 찾기 (대소문자 무시)"""
        target_normalized = target_column.strip().lower()

        for col in columns:
            if col.strip().lower() == target_normalized:
                return row[col]

        return None

    def format_value(self, value) -> str:
        """값의 타입에 따라 적절한 형태로 포맷팅"""
        if pd.isna(value):
            return "NULL"

        # 숫자 타입인지 확인
        if isinstance(value, (int, float)):
            # 정수로 표현 가능한 경우 정수로 변환
            if isinstance(value, float) and value.is_integer():
                return str(int(value))
            return str(value)

        # 문자열인 경우 따옴표로 감싸고 SQL 이스케이핑
        str_value = str(value).replace("'", "''")  # SQL 이스케이핑
        return f"'{str_value}'"

    def write_queries_to_file(self, queries: List[str], output_path: Path):
        """생성된 쿼리들을 텍스트 파일로 저장"""
        try:
            with open(output_path, 'w', encoding='utf-8') as file:
                for i, query in enumerate(queries):
                    file.write(query)
                    if i < len(queries) - 1:
                        file.write('\n')

        except Exception as e:
            raise Exception(f"쿼리 파일 저장 실패: {str(e)}")

    def wait_for_exit(self):
        """사용자 입력 대기 후 종료"""
        print("\n프로그램을 종료하려면 Enter 키를 누르세요...")
        try:
            input()
        except:
            time.sleep(3)  # 3초 후 자동 종료


def main():
    """메인 실행 함수"""
    generator = ExcelQueryGenerator()
    generator.run()


if __name__ == "__main__":
    main()