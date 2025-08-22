# Excel to Query Generator 사용법

## 1. 필요한 라이브러리 설치

```bash
pip install pandas openpyxl
```

## 2. 디렉토리 구조 준비

```
project/
├── excel_query_generator.py
├── sample_query.txt
├── input/
│   ├── 파일1.xlsx
│   ├── 파일2.xlsx
│   └── 파일3.xls
└── output/
    ├── 파일1.txt (생성됨)
    ├── 파일2.txt (생성됨)
    └── 파일3.txt (생성됨)
```

## 3. 쿼리 템플릿 파일 작성

### `sample_query.txt` 파일 예시

**단일 라인 쿼리:**
```sql
update sample_table set use_yn = {사용여부} where sample_pk = {문서번호};
```

**여러 라인 쿼리:**
```sql
update sample_table 
set use_yn = {사용여부}, 
    updated_date = now() 
where sample_pk = {문서번호};
```

## 4. Excel 파일 배치

`input` 디렉토리에 처리할 Excel 파일들을 배치합니다.

## 5. 프로그램 실행

```bash
python excel_query_generator.py
```

## 주요 특징

- ✅ `input` 디렉토리의 모든 Excel 파일(`.xlsx`, `.xls`)을 일괄 처리
- ✅ 각 Excel 파일의 첫 번째 시트만 사용
- ✅ 첫 번째 행을 헤더로 인식
- ✅ 플레이스홀더 `{컬럼명}`을 실제 데이터로 치환
- ✅ 문자열 값은 자동으로 따옴표로 감싸고 SQL 이스케이핑 처리
- ✅ 숫자 값은 따옴표 없이 사용
- ✅ NaN/None 값은 NULL로 처리
- ✅ 대소문자 구분 없이 컬럼명 매칭
- ✅ `output` 디렉토리에 Excel 파일명과 동일한 이름의 `.txt` 파일로 저장
- ✅ 처리 진행 상황과 결과를 콘솔에 출력

## 실행 예시

```
=== Excel to Query Generator ===
input 디렉토리를 생성했습니다.
output 디렉토리를 생성했습니다.
쿼리 템플릿을 읽었습니다: sample_query.txt
템플릿 내용: update sample_table set use_yn = {사용여부} where sample_pk = {문서번호};
발견된 Excel 파일: 2개

처리 중: 사용자데이터.xlsx
  헤더: ['문서번호', '사용여부', '제목']
  데이터 행: 100개
  플레이스홀더: ['사용여부', '문서번호']
  → 100개 쿼리 생성됨
  → 저장 위치: output/사용자데이터.txt
✓ 사용자데이터.xlsx 처리 완료

처리 중: 제품정보.xlsx
  헤더: ['상품ID', '상품명', '가격']
  데이터 행: 50개
  플레이스홀더: ['사용여부', '문서번호']
  경고: '사용여부' 컬럼을 찾을 수 없습니다.
  경고: '문서번호' 컬럼을 찾을 수 없습니다.
  → 50개 쿼리 생성됨
  → 저장 위치: output/제품정보.txt
✓ 제품정보.xlsx 처리 완료

전체 처리 완료: 2/2 파일

처리가 완료되었습니다!
```

## 쿼리 템플릿 예시

### UPDATE 쿼리
```sql
update sample_table set use_yn = {사용여부} where sample_pk = {문서번호};
```

### INSERT 쿼리
```sql
insert into users (name, age, email) values ({이름}, {나이}, {이메일});
```

### DELETE 쿼리
```sql
delete from products where id = {상품ID} and status = {상태};
```

### 복합 UPDATE 쿼리
```sql
update orders 
set status = {주문상태}, 
    updated_by = {수정자}, 
    updated_date = now() 
where order_id = {주문번호} 
  and customer_id = {고객번호};
```

## 주의사항

- Excel 파일의 첫 번째 행은 반드시 헤더(컬럼명)여야 합니다
- 플레이스홀더 `{컬럼명}`의 컬럼명은 Excel 헤더와 정확히 일치해야 합니다 (대소문자 무시)
- 빈 셀이나 NULL 값은 자동으로 `NULL`로 처리됩니다
- 문자열에 포함된 작은따옴표(`'`)는 자동으로 이스케이핑됩니다 (`''`)