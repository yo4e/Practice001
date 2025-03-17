from docx import Document

# 1桁のアラビア数字を単純に漢数字へ置き換えるためのマップ
# （複数桁や特殊な数え方に対応する場合は、さらに処理が必要です）
digit_map = {
    '0': '〇',
    '1': '一',
    '2': '二',
    '3': '三',
    '4': '四',
    '5': '五',
    '6': '六',
    '7': '七',
    '8': '八',
    '9': '九'
}

def convert_digits_to_kanji(text):
    """半角数字を漢数字に変換する（1桁限定の簡易版）。"""
    for d, k in digit_map.items():
        text = text.replace(d, k)
    return text

# Wordファイルを読み込み
doc = Document("/Users/a104/Desktop/input.docx")  # ←ユーザー名に合わせて書き換え

for paragraph in doc.paragraphs:
    # （A）段落の先頭が「(かぎかっこ)で始まるか確認
    if not paragraph.text.startswith("「"):
        # 先頭に全角スペースを1つ挿入（ここを段落インデントにすることも可能）
        paragraph.text = "　" + paragraph.text

    # paragraph.textを直接操作するとRunsが再生成される場合があります。
    # イタリックやボールドなどの情報はRun単位で持っているため、
    # それぞれのRunに対して数値変換＆スタイル変換を行うのが安全です。
    #
    # ただし段落.textを先に書き換えたことでRunsの構造が変わる場合があるため、
    # 一度「先頭一文字下げ」を済ませたあとにRunを操作します。

    for run in paragraph.runs:
        # （B）数字を漢数字に変換
        run.text = convert_digits_to_kanji(run.text)

        # （C）イタリックになっている場合は解除してボールドに
        if run.italic:
            run.italic = False
            run.bold = True

# 処理後のドキュメントを保存
doc.save("/Users/a104/Desktop/output.docx")       # 出力ファイルの保存先も同様に指定
