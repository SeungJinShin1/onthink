from flask import Flask, render_template, request
import pandas as pd
import sys

# 인코딩 문제 해결 (Windows에서 발생할 수 있는 stdout 관련 문제)
sys.stdout.reconfigure(encoding='utf-8')

app = Flask(__name__)

# 엑셀 파일 경로 설정 (book.xlsx 파일로 변경)
excel_file = 'C:/Users/USER/Desktop/python/book.xlsx'

# 엑셀 파일 읽기 (NaN 값을 빈 문자열로 변환)
df = pd.read_excel(excel_file).fillna('')

# 엑셀 데이터 확인 출력 (서버 시작 시 터미널에 출력)
print("엑셀 파일 데이터:")
print(df.head())

# 'Sentence_'로 시작하는 열 이름들만 선택
sentence_columns = [col for col in df.columns if col.startswith('Sentence_')]
print(f"Sentence 관련 열: {sentence_columns}")

# 검색 페이지 (GET 요청 시 index.html로 이동)
@app.route('/')
def index():
    return render_template('index.html')

# 검색 결과 처리 (POST 요청 시 단어 검색)
@app.route('/search', methods=['POST'])
def search():
    # 사용자가 입력한 단어 가져오기
    search_word = request.form.get('word')
    
    # 검색어가 제대로 입력되는지 확인 (터미널에 출력)
    print(f"사용자가 입력한 검색어: {search_word}")
    
    # 검색어가 비어있을 경우 처리
    if search_word is None or search_word.strip() == '':
        return render_template('index.html', message="단어를 입력해 주세요.")

    # 입력된 단어가 엑셀 파일에서 있는지 검색
    results = df[df['Word'].str.contains(search_word, na=False, case=False)]

    # 검색 결과가 있는지 확인
    if not results.empty:
        # 검색 결과 확인 (터미널에 출력)
        print(f"검색 결과: {results}")

        # 'Sentence_'로 시작하는 모든 예시 문장 열을 가져옴
        sentences = results.loc[:, sentence_columns].to_dict(orient='records')
        level = results.iloc[0]['Level']
        return render_template('results.html', word=search_word, level=level, sentences=sentences)
    else:
        # 검색 결과가 없을 경우 처리
        return render_template('index.html', message="단어를 찾을 수 없습니다. 다시 검색해 주세요.")

# Flask 서버 실행 (디버그 모드 비활성화)
if __name__ == '__main__':
    app.run(debug=False)
