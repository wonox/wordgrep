# Wordファイルからのテキスト抽出
from docx import Document
import sys
import os
import re

def wordgrep(path, search_word):
    lis = []
    # document = Document(r'C:\Users\libso23-user\OneDrive - The University of Tokyo\ドキュメント\korekara_doc20221101-1.docx')
    document = Document(path)
    # 1パラグラフごとに配列に入れる
    for i, p in enumerate(document.paragraphs):
        # print(i, p.text)
        i_text = str(i) + ' ｜ ' + p.text
        lis.append(i_text)
    # print(lis)
    # 検索条件
    # pattern = ".*DOI.*"
    pattern = ".*" + search_word +".*"

    # Match the above pattern using list comprehension
    filtered_lis = [x for x in lis if re.match(pattern, x)]

    # 結果の配列の表示
    print(filtered_lis)

    # path_w = 'test_w.txt'
    # with open(path_w, mode='w') as f:
    #     f.write(str(filtered_lis))

# コマンドライン引数を処理する
if __name__ == '__main__':
    args = sys.argv
    if len(sys.argv) < 2:
        sys.exit('Arguments are too short')
    else:
        file_path = sys.argv[1]
        # 絶対パスへ変換
        path = os.path.abspath(file_path)
        # ファイルの存在有無
        if not os.path.isfile(path):
            sys.exit("I don't have that file.")
        # .docxを想定している場合
        if path.split(".")[-1] != "docx":
            sys.exit("That's not the right extension!!!")
        search_word = sys.argv[2]
        wordgrep(path, search_word)


# python wordgrep.py "C:\Users\libso23-user\OneDrive - The University of Tokyo\ドキュメント\20221031要求水準書ひな形案.docx" 臭い


