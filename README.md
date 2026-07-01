"# Automated-consolidation-of-cash-flow-data_SHUC" 
# 현금흐름 엑셀 데이터 취합 자동화 매뉴얼

<aside>
💡

1. 목적 : 부서 별 현금 흐름 엑셀 데이터를 **자동 취합**하기 위함.
2. 프로그램 작동 흐름
3. 설정 사항 (코드에서 변경 가능)
    1. 날짜 계산을 자동으로 하도록 함.
        - 오늘이 2026.06.22 → 다음달(26.07) 6월 분 현금 흐름 엑셀 데이터 작성
            
            프로그램이 자동 날짜 계산하여 파일(ex. 도시개발본부 현금흐름 분석자료(26.06))을 탐색함
            
    2. 폴더/파일명
        - 부서별 현금흐름 파일들을 저장한 파일명을 ‘부서명파일’로 함.
        - 취합본파일 명을 ‘도시개발본부_현금흐름_최종취합_완성본({prev_year_short}.{prev_month:02d}).xlsx’로 함.
    3. 인덱스열은 H열로 고정.  
    4. 각종 사업의 종류는 A~Z 알파벳 코드 사용함.
    5. 도시개발본부 부서의 코드명은 다음과 같이 설정함. 
        
        
        | 부서명 | 코드명 | 뜻 |
        | --- | --- | --- |
        | 도시개발본부 | UDH | Urban Development Headquarter |
        | 도시기획부 | UP | Urban Project |
        | 개발사업1부 | DB1 | Development Business 1 |
        | 개발사업2부 | DB2 | Development Business 2 |
        | 관광사업부 | TB | Tourism Business  |
        | 도시개발1부 | UD1 | Urban Development 1 |
        | 도시개발2부 | UD2 | Urban Development 2 |
        | 도시사업1부 | UB1 | Urban Business 1 |
        | 도시사업2부 | UB2 | Urban Business 2 |
        | 법인세 외 | CT | Cooperate Tax |
        - 부서가 추가·변경되면 코드 내 변수에 추가 (utils.py)
            - `dept_index_codes` 와 `PREFIX_I_LABELS` 두 곳만 맞추면 됩니다.
    6. 색상
        1. 부서별로 수정된 셀 → 취합본 파일의 수정된 셀에 파란색 으로 표시 
            
            <img width="229" height="150" alt="Image" src="https://github.com/user-attachments/assets/0865d094-f9ba-44e9-9e99-e61b18937c50" />
            
        2. 부서별로 행을 추가한 경우 → 취합본 파일의 인덱스 열에 연두색으로 표시
            
            <img width="225" height="144" alt="Image" src="https://github.com/user-attachments/assets/40f9ed57-70b0-4740-a80b-20a16c671c2b" />
            
        3. 부서별로 행을 삭제한 경우 → 취합본 파일의 인덱스 열에 연한 빨간색으로 표시
            
            <img width="219" height="146" alt="Image" src="https://github.com/user-attachments/assets/1e6c493f-36c5-4bae-b9bd-b797717550e8" />
            
4. 9개 파일+ 시트 7개 (총 63개 sheet) 취합에 걸린 총 시간 : (2026-07-01 기준)
    
    소요 시간: 2분 21초
    최종 취합본 파일 저장완료
    -> 경로: C:\Users\user\Desktop\인턴\3)현금흐름\도시개발본부_현금흐름_최종취합_완성본(26.05).xlsx
    
</aside>

---

## 1️⃣ 환경 설정

1. python 3.12 다운로드
    
    https://www.python.org/downloads/release/python-3120/
    
    <img width="1261" height="645" alt="Image" src="https://github.com/user-attachments/assets/3f682496-f62b-4774-b1d1-a05b8d7349a5" />
    
    - Set-up
        - Path 추가
            
            <img width="650" height="403" alt="Image" src="https://github.com/user-attachments/assets/3aaa2d0b-183b-4d59-8547-12baa6ee35be" />
            
            - **❗주의 :  반드시 Add python.exe to PATH 누르기!!**
            - IDLE 자동 설치
2. Package 설치
    
    <img width="774" height="675" alt="Image" src="https://github.com/user-attachments/assets/b1cd8e37-19c2-4cae-9ed7-0ede38c0a423" />
    
    - cmd 실행
    - 명령문 입력
        
        ```python
        pip install openpyxl    # 숫자, 텍스트 취합할 때 사용할 라이브러리
        pip install pywin32     # 대시보드1, 대시보드2 취합할 때 사용할 라이브러리
        ```
        
        - 실행 결과
            
            <img width="1092" height="346" alt="Image" src="https://github.com/user-attachments/assets/7b78e5ff-8942-46f7-8d62-ae33aef0fbb1" />
            
            - [notice] To update, run : python.exe -m pip install —upgrade pip 가 뜨면, python.exe -m pip install —upgrade pip 도 실행
            (Ctrl+Shift+C, Ctrl+Shift+V 복사 붙여넣기 가능)
            
3. 부서별 파일 같은 경로에 저장
    
    <img width="1006" height="390" alt="Image" src="https://github.com/user-attachments/assets/e3c8ec04-3563-49fd-b074-6efb8764e36b" />
    
    예시 폴더
    

---

## 2️⃣ ⭐ 전체 흐름 요약 ⭐

### 1. 취합용 파일 양식 복사

- 취합 시에도 레이아웃 등을 유지하기 위함
- 복사 대상
    - `IndexStore`
        
        ```python
        { 'U-4' : {
        							'row_num' : 370,  # 행번호
        							'value' : {
        													9: '사업명', 10: '구분', .... , 14 : 140,000
        													}     # 열 번호 : 값 
        					}
        
        -> 인덱스의 행번호와 열번호에 따른 값을 저장
        ```
        
        - 필요한 이유 : 행을 삽입하면 아래 행 전체가 밀림 → 취합 시, 텍스트와 값이 깨질 수 있음 → IndexStore의 ‘삽입 전 원본값’을 밀린 행에 덮어씀.
    - `Section baseline`
        
        ```python
        { 'U' : {
        					'start_row' : 350,     # 350 행 부터 U 관련 내용
        					'rel_end' : 30,        # 350 ~ 380 범위까지 U 관련 내용
        					'row' : { 5 : {9: '지출', ...., 13 : 999,...},
        									} # 355 행 의 9열은 '지출', 13열은 999 값을 가짐
        				 }
         }
         
         -> 부서 인덱스 시작 행, 섹션 범위, 행과 열 값을 저장
        ```
        
    - `original master layout`
        
        ```python
        { 'U' : { 'rel_end_row' : 30,
        					'merges' : [ 'rel_min_row' : 0, 'rel_max_row' : 0,
        												'min_col' : 9, 'max_col' :9,
        												'value' : 'U', 'style' : {fill, font, border}
        											]
        					'row_cells'
        					}
        }
        -> 병합범위(merges)와 값/테두리 색이 있는 셀의 값과 서식을 저장			
        ```
        

### 2. 취합용 파일, 부서별 파일 로드

- 취합용 파일 1번 / 부서별로 파일 1번(부서가 N개 있으면 파일 N번 로드) 로드하여 속도 개선
    
    (시트 마다 파일을 재로드하지 않음) 
    

### 3. 숫자 시트 취합

- 인덱스 기반 탐색
    
    
    |  | 취합용 파일 | 부서 파일 | 처리 결과 |
    | --- | --- | --- | --- |
    | 수정된 셀이 포함된 행의 인덱스 | 존재 | 존재 | 부서에서 수정한 값과 **파란색 배경색**을 취합용 파일에 덮어씀 |
    | 추가된 행의 인덱스 | 존재X | 존재
    **(부서에서 직접 입력해야함)** | 취합용 파일에 행을 추가하고 신규 인덱스 셀에 **연두색 배경색**으로 덮음 |
    | 삭제된 행의 인덱스 | 존재 | 존재X | 취합용 파일에 인덱스 셀에 **연한 빨간색 배경색**을 덮고, 행의 데이터는 삭제하지 않음 |
- 탐색 방법
    - dept_map : 인덱스와 행번호를 저장 = H열 스캔값
        
        ```python
        {'U-4' : {'row_num': 370}}
        ```
        
- 부서에서 **셀을 수정**한 경우 :
    - 취합시 파란셀 셀이 포함된 행을 전체 복사함.
- 부서에서 **행을 추가**한 경우 : ⭐인덱스열(H열)의 셀값을 **직접 입력**해야 함⭐
    - 인덱스열(H열)의 **셀값을 직접 입력**
        - 삭제행과 구분하기 위함
        - 입력 방법 :
            - [ 예시 ] 만약 도시기획부UP 에 해당하는 인덱스열(H열)의 범위가 UP-1~UP-9라고 가정도시기획부에서 UP-4와 UP-5 사이에 행을 추가 → 추가된 행의 인덱스열(H열)에는 UP-10를 입력해야함
                
                
                | UP-1 |
                | --- |
                | UP-2 |
                | UP_3 |
                | UP-4 |
                | UP-5 |
                | UP-6 |
                | UP-7 |
                | UP-8 |
                | UP-9 |
                
                **[기존의 인덱스 열]**
                
                | UP-1 |
                | --- |
                | UP-2 |
                | UP_3 |
                | UP-4 |
                | **UP-10** |
                | UP-5 |
                | UP-6 |
                | UP-7 |
                | UP-8 |
                | UP-9 |
                
                **[행 추가 후 인덱스 열]**
                
            - 만약 부서에서 인덱스 값을 잘못 기입한 경우, **자동보정** 되도록 설정함
                
                [예시] 도시기획부에서 UP-10을 UP-5로 입력하고, 그 이후의 행에 대해서 UP-6~UP-10으로 잘못 입력한 경우 → 코드에서 자동보정되도록 설정해두었음
                
    - 행 추가 시, **함수 수식 내 행 번호도 갱신**
        - 함수 수식 내 행번호를 갱신하지 않으면 취합시, 잘못된 함수결과 값이나 값이 사라질 위험(#REF!, #N/A 등 오류 발생) 이 있기 때문
        - 갱신 과정
            1. 취합용 파일 → 원본, 수식을 보관 (IndexStore, Section Baseline)
            2. 부서에서 행 추가 → insert_row 확인
            3. baseline으로 원본, 수식 복원 (이때, 행번호는 행 추가 이전의 번호임)
            4. 레이아웃 복원
            5. fix_formula_after_inserts() 함수를 사용해 함수수식 내 행번호 갱신
                - 2번의 insert_row 개수에 따라 함수 수식을 갱신
    - 취합 시, 취합용 파일에 없는 H열 인덱스값을 발견하면, new_rows_buffer에 적재 (대기열의 개념)
- 부서에서 **행을 삭제**하고 싶은 경우 : ⭐인덱스열(H열)의 셀값**만** 지우면 됨 ⭐
    - 삭제의 경우에는, 관리자의 확인이 필요하다고 생각되어 부서에서 행 전체를 삭제하지 않고 인덱스열(H열)에 해당하는 셀값만 삭제 → **인덱스열(H) 의 셀값만 빈칸으로 둠**
        - [예시] 도시기획부에서 UP-1 인덱스에 해당하는 행을 삭제하고 싶다면, UP-1 값만 빈칸으로 둠 + 빈칸은 연한 빨간색으로 칠함 → 취합 후, 관리자가 UP-1에 해당하는 행들을 확인 → 관리자의 삭제

### 4. 레이아웃 복원

- 행을 추가/삭제하다보면, 병합셀이 깨지거나 레이아웃이 무너짐 →  취합용 파일 양식을 복사한 것을 토대로 레이아웃 복원
- 레이아웃 복원 과정
    1. 행 추가 전, original master layout을 섹션별 rel 좌표로 다시 깔음
    2. J열의 수입/지출/잔액/발생이자 ‘항목’열에 대해 세로 병합 재생성
    3. 신규행 : remap 
    

### 5. 텍스트 시트 취합

- 텍스트 시트에는 인덱스 번호가 없음 → 한 줄씩 검사하다가 파란색 배경색을 가진 셀을 발견하면(수정된 셀) 그 행 전체를 취합용에 복사

### 6. 대시보드 시트 취합

- 대시보드 시트는 도형으로 인식시켜서 텍스트 추출함. → **도형 간 그룹화를 해제해야 인식 가능**
- 대시보드 시트의 경우, 텍스트 상자의 위치를 자유롭게 배치하기 때문에, 텍스트 삽입 자동화가 어려움
- 따라서, 별도의 대시보드 텍스트 취합용xlsx 파일을 만들어 부서에서 추가한 내용을 **빨간색 배경색**으로 표시함 → 관리자가 확인 후, 대시보드 차트 내에 추가
    
    <img width="905" height="53" alt="Image" src="https://github.com/user-attachments/assets/e1a857da-61c6-4b7e-8ca1-6b3fe5acc7c8" />
    
    대시보드 취합본 엑셀 파일
    
  <img width="1707" height="768" alt="Image" src="https://github.com/user-attachments/assets/ec2c53f5-4110-4d31-84ab-0f245da020bb" />
    
    대시보드 취합본 엑셀 파일 - 부서에서 추가한 내용은 빨간색 배경으로 표시
    

### 7. 최종 파일 저장

---

## 3️⃣ Python 파일 내 함수 분류

> **utils.py** → 유틸리티 함수들을 모아둔 파일
**main.py** → 숫자/텍스트 시트를 `취합을 실행`하는 파일
> 

→ 함수 별 자세한 작동 흐름은 **utils.py, main.py 내 주석**으로 추가

### 1. 유틸리티 함수

1. 날짜, 환경설정 관련
2. 셀/서식 관련
3. 색상 판별 관련
4. 인덱스 관련
5. 부서명 라벨 관련
6. 병합셀/레이아웃 관련
7. 파일 입출력 관련

### 2. 숫자 sheet  (원안, 수정1안, 수정2안) 함수

1. aggregate_numeric_sheets 
2. 실행 결과 : 
    
    <img width="497" height="559" alt="Image" src="https://github.com/user-attachments/assets/23991a43-a4e0-49d7-849d-fb546a3de42b" />
    
   <img width="633" height="547" alt="Image" src="https://github.com/user-attachments/assets/44ceb6df-507f-4ee4-a958-31e7b9ff6635" />
    
    - 시트 끼리(취합본 파일의 시트 + 부서별 파일의 시트), **인덱스 기반**으로 수정된 셀을 찾아 수정된 값과 파란색 배경색으로 대체 / 행 추가 및 삭제 + 레이아웃 셀 서식 유지
    

### 3. 텍스트 sheet  (전월대비 비교, 지출계획) 함수

1. aggregate_text_sheets
2. 실행 결과
    
   <img width="584" height="230" alt="Image" src="https://github.com/user-attachments/assets/fcdc068c-651c-427d-9438-e15511a8a0e0" />
    
    - 시트끼리 (취합본 파일의 시트 + 부서별 파일의 시트),  **행 탐색 기반**으로 수정된 셀을 찾아 수정된 값과 파란색 배경색으로 대체
    

### 4. 대시보드 sheet (DashBoard1, DashBoard2) 함수

1. 대시보드 찾는 함수
2. 전처리 + 정렬 함수 (’`’, ‘(’ 를 삭제하고 오름차순으로 정렬)
3. 부서에서 추가한 내용을 취합하는 함수
4. 대시보드 내 부서에서 추가한 내용을 취합하여 저장하는 함수
5. 실행 결과
    
    !image.png
    

---

## 4️⃣ Python 실행

### 1. 실행 전, utils.py 에서 폴더경로와 파일명 확인

!image.png

- base_path : 이 경로에 부서별폴더과 파이썬파일(utils.py, main.py), 취합용 파일이 모두 있어야 함.
- file_name : 취합용 파일
    - 이미지 내 prev_year_short, prev_month:02d 는 날짜 자동으로 계산하기 위한 변수
- 최종 파일명 : main.py 코드 내 취합결과 최종 파일명 지정함.
    
    !image.png
    
    - ex.도시개발본부_현금흐름_최종취합(26.05).xlsx

### 2. Python 파일(main.py) 실행

- IDLE 실행 →  File → Open → **main.py 파일** 클릭 → **F5 단축키** 누르기
    
    !image.png
    
    !image.png
    
    - **main.py** : utils.py , dashboard.py 내 함수들을 불러와서 숫자 시트, 텍스트 시트, 대시보드 시트 취합을 한번에 **실행**하는 파일

### 3. 최종 취합본 파일 생성 및 저장

!image.png

- 최종 취합본 파일은 base_path에 저장
- 최종 취합본 파일명 : 도시개발본부_현금흐름_최종취합_완성본({연도}.{전월}).xlsx
(예시: 도시개발본부_현금흐름_최종취합_완성본(26.05).xlsx)
