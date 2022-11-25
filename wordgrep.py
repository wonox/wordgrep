# Text extraction from Word(docx) files
from docx import Document
import sys
import argparse
import os
import re

# outline --help
parser = argparse.ArgumentParser(
    description="Text extraction from Word(docx) files")

parser.add_argument(
    '-f',
    '--file',
    type=str,
    help='path\\filename.docx',
    required=True)

parser.add_argument(
    '-w',
    '--word',
    type=str,
    help='search word(Regular expression enabled)',
    required=True)

cmd_args = parser.parse_args()
file_path = cmd_args.file


def wordgrep(path, search_word):
    lis = []
    document = Document(path)
    # 1パラグラフごとに配列に入れる
    for i, p in enumerate(document.paragraphs):
        i_text = str(i) + ' : ' + p.text
        lis.append(i_text)
    # 検索条件
    pattern = ".*" + search_word + ".*"

    # Match the above pattern using list comprehension
    # filtered_lis = [x for x in lis if re.match(pattern, x)]
    filtered_lis = [x for x in lis if re.fullmatch(pattern, x, re.IGNORECASE)]

    # 結果の配列の表示
    # print(filtered_lis)
    print('\n'.join(filtered_lis))


# コマンドライン引数を処理する
if __name__ == '__main__':

    if not file_path:
        sys.exit('Arguments are too short')
    else:
        # 絶対パスへ変換
        path = os.path.abspath(file_path)
        # ファイルの存在有無
        if not os.path.isfile(path):
            sys.exit("I don't have that file.")
        # .docxを想定している場合
        if path.split(".")[-1] != "docx":
            sys.exit("That's not the right extension!!!")
        search_word = cmd_args.word
        wordgrep(path, search_word)
