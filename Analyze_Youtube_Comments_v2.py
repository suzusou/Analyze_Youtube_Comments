import googleapiclient.discovery
import requests
import json
import openpyxl
from openpyxl import Workbook
import matplotlib.pyplot as plt
from tkinter import simpledialog

# YouTubeのAPIキーを入力してもらう関数
def get_api_key():
    api_key = simpledialog.askstring("YouTube APIキー", "YouTube Data API v3のAPIキーを入力してください:")
    # キャンセル又は値が空の場合、プログラム終了
    if api_key is None or api_key.strip() == "":
        print("APIキーが入力されていません。")
        exit()
    return api_key

# Video IDを入力してもらう関数
# 動画IDを指定してコメントを取得します
# 例)　https://www.youtube.com/watch?v=yeZ3STy3k44
# v=　以降の値をvideo_idに記述
def get_video_id():
    video_id = simpledialog.askstring("YouTube Video ID", "YouTubeのVideo IDを入力してください:")
    # キャンセル又は値が空の場合、プログラム終了
    if video_id is None or video_id.strip() == "":
        print("Video IDが入力されていません。")
        exit()
    return video_id

# COTOHAのクライアントIDとクライアントシークレットを入力してもらう関数
def get_cotoha_credentials():
    client_id = simpledialog.askstring("COTOHA クライアントID", "COTOHAのデベロッパークライアントIDを入力してください:")
    client_secret = simpledialog.askstring("COTOHA クライアントシークレット", "COTOHAのデベロッパークライアントシークレットを入力してください:")
    # キャンセル又は値が空の場合、プログラム終了
    if client_id is None or client_id.strip() == "" or client_secret is None or client_secret.strip() == "":
        print("COTOHAのクライアント情報が入力されていません。")
        exit()
    return client_id, client_secret

# YouTubeのコメントを取得する関数
def get_youtube_comments(api_key,video_id):
    comments = []
    try:
        # YouTube Data API v3との接続を確立
        youtube = googleapiclient.discovery.build('youtube', 'v3', developerKey=api_key)

        # APIの都合で１００件までのコメントを取得
        response = youtube.commentThreads().list(
            part='snippet',
            videoId=video_id,
            textFormat='plainText',
            maxResults=100
        ).execute()

    except googleapiclient.errors.HttpError:
        print("APIキー又はVideo IDが間違っています。")
        exit()

    # コメントをdataリストに格納
    for item in response['items']:
        comment = item['snippet']['topLevelComment']['snippet']['textDisplay']
        comments.append(comment)

    return comments


# COTOHAのAPIは無料プランで1日1000件まで可能
def get_cotoha_sentiment(text, client_id, client_secret):
    
    # Developer API Base URL
    base_url = 'https://api.ce-cotoha.com/api/dev/nlp/v1/sentiment'

    # headerの情報を記述
    headers = {
        'Content-Type': 'application/json',
        'Authorization': f'Bearer {get_access_token(client_id, client_secret)}'
    }
    
    # dataにYoutubeのコメントを格納
    data = {
        'sentence': text
    }

    # COTOHAの感情分析のAPIに向かって、POSTリクエストを送信
    response = requests.post(base_url, headers=headers, data=json.dumps(data))

    # JSON形式のデータにして格納
    result = response.json()

    return result

# COTOHAに対して、認証情報を送信する
def get_access_token(client_id, client_secret):

    url = 'https://api.ce-cotoha.com/v1/oauth/accesstokens'

    headers = {
        'Content-Type': 'application/json'
    }
    
    data = {
        'grantType': 'client_credentials',
        'clientId': client_id,
        'clientSecret': client_secret
    }

    response = requests.post(url, headers=headers, data=json.dumps(data))
    result = response.json()

    # accessトークンを返す
    try:
        return result['access_token']
    except KeyError :
        print("COTOHAのクライアント情報が間違っています。")
        exit()

def main():
    # Tkinterのダイアログを使用してAPIキー、Video ID、COTOHAのクライアント情報を取得
    api_key = get_api_key()
    video_id = get_video_id()
    client_id, client_secret = get_cotoha_credentials()

    # 日本語フォントの設定
    plt.rcParams["font.family"] = "Meiryo"

    # Excelファイルを新規作成
    workbook = Workbook()
    sheet = workbook.active

    # dataの列部分を定義
    data = [["No.", "コメント","感情"]]
    
    # Youtubeのコメント欄を取得
    comments = get_youtube_comments(api_key,video_id)
    emotion=""

    # Youtubeのコメント分以下の処理を実行
    for index, comment in enumerate(comments, 1):

        # Youtubeのコメント欄をNTT docomoさんのCOTOHAの感情分析のAPIにかける
        result = get_cotoha_sentiment(comment,client_id, client_secret)

        # 結果に応じて、emotionに値を格納
        if result['result']['sentiment']=='Positive':
            emotion="ポジディブ"
        elif result['result']['sentiment']=='Negative':
            emotion="ネガティブ"
        else:
            emotion="ニュートラル"

        # dataに処理結果を格納    
        data.append([index,comment,emotion])

    # Excelファイルにdataを格納
    for row in data:
        sheet.append(row)
    
    # Excelファイルを保存
    workbook.save("youtube_comments.xlsx")

    # ネガティブ、ポジティブ、ニュートラルの数を数える
    num_negative = sum(1 for row in data[1:] if row[2] == "ネガティブ")
    num_positive = sum(1 for row in data[1:] if row[2] == "ポジディブ")
    num_neutral = sum(1 for row in data[1:] if row[2] == "ニュートラル")

    # 円グラフ用のデータを準備
    labels = ["ネガティブ", "ポジディブ","ニュートラル"]
    values = [num_negative, num_positive,num_neutral]

    # 円グラフを作成
    plt.figure(figsize=(6, 6))
    plt.pie(values, labels=labels, autopct="%1.1f%%", startangle=90)
    plt.title("Youtubeのコメント欄の批判コメントの割合")
    plt.axis("equal")  # アスペクト比を等しくすることで円形に

    # Excelファイルにグラフを挿入する位置を指定（この例ではセルB10から）
    chart_cell = "E1"

    # グラフを画像として保存
    chart_image = "youtube_comments.png"
    plt.savefig(chart_image, dpi=50)

    # 画像をExcelファイルに挿入
    img = openpyxl.drawing.image.Image(chart_image)
    img.anchor = chart_cell
    sheet.add_image(img)

    # Excelファイルを上書き保存
    workbook.save("youtube_comments.xlsx")


if __name__ == "__main__":
    main()
