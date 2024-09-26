import os
import re
import discord
import requests
from google.cloud import vision
from discord.ext import commands
from discord import app_commands
import gspread
from google.oauth2.service_account import Credentials
from difflib import SequenceMatcher 
import pytz
from datetime import datetime
import time
from typing import Optional

# JSTタイムゾーン
jst = pytz.timezone('Asia/Tokyo')

# Google Cloud Vision APIとGoogle Sheets APIのクレデンシャルファイルのパス
json_file_path = ''#適切に指定してください

# Google Cloud Vision APIのクレデンシャルファイルの指定
os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = json_file_path

# Google Vision APIクライアントを作成
vision_client = vision.ImageAnnotatorClient()

# Discord Botのトークンを設定
TOKEN = ''#適切に指定してください

# 有効なチャンネルを保存するリスト
active_channels = []

# グローバル変数の宣言とデフォルトのスプレッドシートURL
sheet_url = "" #適切に指定してください

# Intentsの設定
intents = discord.Intents.default()
intents.message_content = True
intents.members = True  # メンバー情報を取得するためのIntentを有効にする


# Discordクライアントの作成
class MyBot(discord.Client):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.tree = app_commands.CommandTree(self)

    async def on_ready(self):
        await self.tree.sync()  # コマンドをDiscordに同期
        print(f'Logged in as {self.user}!')

# クライアントの初期化
bot = MyBot(intents=intents)

# グローバル変数の宣言
start_dt = None
end_dt = None

@bot.tree.command(name="yomitorikun", description="リザルト読み取りのオン/オフを切り替え、画像解析を行います")
async def yomitorikun(interaction: discord.Interaction, 
                      on_off: bool = True,  # 読み取りのオン/オフ (デフォルトでTrue)
                      sheet: Optional[str] = None,  # シートURLをオプションで受け取れるようにする
                      start_year: Optional[int] = None, start_month: Optional[int] = None, start_day: Optional[int] = None, 
                      start_hour: Optional[int] = None, start_minute: Optional[int] = None, 
                      end_year: Optional[int] = None, end_month: Optional[int] = None, end_day: Optional[int] = None, 
                      end_hour: Optional[int] = None, end_minute: Optional[int] = None):
    global sheet_url, start_dt, end_dt, active_channels

    # 読み取りをオフにする場合
    if not on_off:
        channel_id = interaction.channel_id
        if channel_id in active_channels:
            active_channels.remove(channel_id)
            await interaction.response.send_message(f"リザルト読み取り機能を無効にしました。")
        else:
            await interaction.response.send_message(f"このチャンネルはすでに無効化されています。")
        return

    # シートのURLを受け取った場合は上書き、受け取らなければデフォルトを使用
    if sheet:
        sheet_url = sheet
        await interaction.response.send_message(f"Googleスプレッドシートが設定されました: {sheet_url}")
    else:
        await interaction.response.send_message(f"スプレッドシートのURLが未指定だったため、デフォルトURLを使用しました")

    # チャンネルの画像解析を有効にする
    channel_id = interaction.channel_id
    if channel_id not in active_channels:
        active_channels.append(channel_id)
        await interaction.followup.send("リザルト読み込み機能が有効になりました。")
    else:
        await interaction.followup.send(f"{interaction.channel.name} チャンネルはすでにアクティブです。")
        return

@bot.tree.command(name="create_thread", description="チームごとにプライベートスレッドを作成します")
async def create_thread(interaction: discord.Interaction, team_id: Optional[int] = None):
    await interaction.response.defer()

    # デフォルトのシートURLを設定
    global sheet_url
    if not sheet_url:
        sheet_url = "https://docs.google.com/spreadsheets/d/1qAbN-OMvG9flO7ey9m5ylHg0-k0qkHc1zUrnSPbHtqg/edit#gid=991675042"
    
    # スプレッドシートIDを抽出して有効なURLか確認
    try:
        team_list = get_team_list(sheet_url)
    except gspread.exceptions.NoValidUrlKeyFound:
        await interaction.followup.send("指定されたスプレッドシートのURLが無効です。確認してください。")
        return

    guild_members = interaction.guild.members

    if team_id is not None:
        team = next((team for team in team_list if team['No'] == team_id), None)
        if team:
            try:
                await create_team_thread(interaction, team, guild_members)
                await interaction.followup.send(f"チームID {team_id} のスレッドが作成されました。")
            except Exception as e:
                await interaction.followup.send(f"チームID {team_id} のスレッド作成中にエラーが発生しました: {e}")
        else:
            await interaction.followup.send(f"指定されたチームID {team_id} が見つかりませんでした。")
    else:
        failed_teams = []
        for team in team_list:
            try:
                await create_team_thread(interaction, team, guild_members)
                # スレッド作成後に少し待機してレートリミットを回避
                time.sleep(1.5)  # 1.5秒の待機（必要に応じて調整可能）
            except Exception as e:
                failed_teams.append(f"{team['No']}: {e}")
                print(f"Error creating thread for team No {team['No']} - {team['チーム名']}: {e}")

        if failed_teams:
            error_message = "\n".join(failed_teams)
            await interaction.followup.send(f"以下のチームのスレッド作成中にエラーが発生しました:\n{error_message}")
        else:
            await interaction.followup.send("すべてのチームのスレッドが作成されました。")


# チームごとのスレッドを作成する関数
async def create_team_thread(interaction, team, guild_members):
    team_no = team['No']  # チームのNoを取得
    team_name = team['チーム名']
    discord_users = team['Discordユーザ名']  # Discordユーザ名のリストを取得

    # スレッドタイトルを生成
    timestamp = time.strftime("%Y%m%d-%H%M%S")
    thread_name = f"{team_no}_{team_name}_{timestamp}"  # No_チーム名_タイムスタンプ の形式

    # プライベートスレッドを作成
    thread = await interaction.channel.create_thread(name=thread_name, type=discord.ChannelType.private_thread)

    # ユーザーをスレッドに追加
    if not discord_users:
        # Discordユーザ名がない場合のメッセージ
        await thread.send("Discordユーザ名の記載がなかったため、該当のユーザを追加できませんでした。")
    else:
        # カンマ区切りで複数のDiscordユーザ名を分割
        discord_users_list = discord_users.split(',')
        for user_name in discord_users_list:
            user_name = user_name.strip()  # 前後のスペースを削除

            # サーバー内メンバーからユーザ名またはディスプレイネームで一致するユーザーを検索
            user = next((member for member in guild_members if member.name.lower() == user_name.lower() or member.display_name.lower() == user_name.lower()), None)

            if user:
                # スレッドにユーザーを追加
                await thread.add_user(user)
            else:
                # ユーザーが見つからなかった場合のメッセージ
                await thread.send(f"ユーザ {user_name} が見つかりませんでした。")

    # 運営ロールをsilent=Trueでメンション
    role = discord.utils.get(interaction.guild.roles, name="運営")  # 運営ロールの取得
    if role:
        await thread.send(content=f"{role.mention}", silent=True)  # 運営ロールをメンション
    
    # 新しいメッセージをスレッドに送信
    await thread.send(f"「{team_name}」チーム専用の画像提出用スレッドです。必要に応じてメンバーをメンション（@ユーザ名）で追加してください。\n"
                      "また、画像は１つのすてっぷごとに提出してください（まとめて送付しても処理できません）")


@bot.tree.command(name="delete_all_threads", description="このチャンネル配下のすべてのスレッドを削除します")
async def delete_all_threads(interaction: discord.Interaction, password: str):
    """パスワードが正しい場合のみ、このチャンネルに関連するすべてのスレッドを削除する"""
    
    # パスワードの確認
    if password != "yuzushikakatan":
        await interaction.response.send_message("パスワードが正しくありません。")
        return
    
    # 処理を遅らせてタイムアウトを回避
    await interaction.response.defer()

    # コマンドが実行されたチャンネル
    channel = interaction.channel

    # スレッド一覧を取得し、削除
    for thread in channel.threads:
        try:
            await thread.delete()  # スレッドを削除
        except Exception as e:
            team_no = thread.name.split('_')[0]
            print(f"Failed to delete thread for チームNo {team_no}: {e}")

    # すべてのスレッドが削除されたことを通知
    await interaction.followup.send("すべてのスレッドを削除しました。")

@bot.tree.command(name="check_id", description="Discord IDの実在確認を行います")
async def check_id(interaction: discord.Interaction):
    global sheet_url
    
    # スプレッドシートからチーム名リストを取得
    try:
        team_list = get_team_list(sheet_url)
    except gspread.exceptions.NoValidUrlKeyFound:
        await interaction.response.send_message("指定されたスプレッドシートのURLが無効です。確認してください。")
        return

    # メンションが動作していないDiscordユーザを確認するメッセージ
    message_text = (
        "本大会では、結果報告の際にDiscordIDが必要なため、定期的にDiscord IDの実在確認を行っています。\n"
        "ご自身のチームで「ユーザ名が見つかりませんでした」と表示される場合は、\n"
        "Discord IDを再度ご確認いただき、適切なIDをお知らせください。\n"
    )
    
    # チャットメッセージを送信
    await interaction.response.send_message(message_text)

    # 現在の日時をJSTに変換してスレッドタイトルを生成
    jst = pytz.timezone('Asia/Tokyo')  # JSTタイムゾーンの取得
    current_time_utc = datetime.now(pytz.utc)  # 現在のUTC時間を取得
    current_time_jst = current_time_utc.astimezone(jst)  # UTCをJSTに変換
    thread_title = current_time_jst.strftime("%Y%m%d_Discord ID申請状況")  # 日付のみでスレッドタイトルを生成

    # スレッドを作成
    thread = await interaction.channel.create_thread(name=thread_title, type=discord.ChannelType.public_thread)

    # サーバー内の全メンバーを取得
    guild_members = interaction.guild.members

    # スレッドの内容を構築
    thread_content = ""
    for team in team_list:
        team_no = team['No']
        team_name = team['チーム名']
        discord_users = team['Discordユーザ名'].split(',')
        
        # Discordユーザ名に基づいてメンション形式に変換
        mentions = []
        for user_name in discord_users:
            user_name = user_name.strip()

            # サーバー内メンバーからユーザ名またはディスプレイネームで一致するユーザーを検索
            user = next((member for member in guild_members if member.name.lower() == user_name.lower() or member.display_name.lower() == user_name.lower()), None)

            if user:
                # 一致するユーザーをメンションに追加
                mentions.append(user.mention)
            else:
                # 一致するユーザーが見つからなかった場合
                mentions.append(f"@{user_name}（ユーザが見つかりませんでした）")

        mentions_str = ', '.join(mentions)

        # スレッドに追加するチーム情報
        team_info = f"No {team_no}: {team_name} - {mentions_str}\n"
        
        # スレッド内容が2000文字を超えないようにメッセージを分割して送信
        if len(thread_content) + len(team_info) > 2000:
            await thread.send(thread_content)  # 現在の内容を送信
            thread_content = team_info  # 新しいメッセージとして再スタート
        else:
            thread_content += team_info  # 現在の内容に追加

    # 最後に残ったメッセージを送信
    if thread_content:
        await thread.send(thread_content)




# 画像からテキストを抽出する関数
def detect_text(image_path):
    """Google Cloud Vision APIを使用して画像からテキストを抽出する"""
    with open(image_path, 'rb') as image_file:
        content = image_file.read()

    image = vision.Image(content=content)
    response = vision_client.text_detection(image=image)
    texts = response.text_annotations

    if response.error.message:
        raise Exception(f'{response.error.message}')
    
    # description全体を出力
    all_texts = '\n'.join([text.description for text in texts])
    return texts, all_texts  # texts と all_texts を返す

def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()

def extract_stage_name(texts):
    """ステージ名が検出された文字列に含まれるかどうかを確認してステージ名を抽出する"""
    stages = [
        "アラマキ砦", "ムニ・エール海洋発電所", "シェケナダム", 
        "難破船ドン・ブラコ", "すじこジャンクション跡", 
        "トキシラズいぶし工房", "どんぴこ闘技場", "グランドバンカラアリーナ"
    ]
    
    # 検出時に使用するステージ名の一部
    stages_2 = [
        "アラマキ", "発電所", "シェケナダム", 
        "ブラコ", "すじこジャンクション", 
        "トキシラズ", "闘技場", "グランドバンカラアリーナ"
    ]
    
    for text in texts:
        description = text.description.strip()
        for partial_stage, full_stage in zip(stages_2, stages):
            # 検出された文字列にステージ名の一部が含まれるかを確認
            if partial_stage in description:
                return full_stage  # 一致した場合は元のフルステージ名を返す
    return "ステージ名が見つかりません"

# 特定の情報を抽出する汎用関数
def extract_specific_info(texts, pattern):
    """特定の情報（日付、クリア状況、キケン度など）を抽出する"""
    result = []
    for text in texts:
        description = text.description.strip()
        match = pattern.search(description)
        if match:
            result.append(match.group())
    # 重複を排除してリストに戻す
    return list(set(result))

# WAVE情報を抽出する関数
def extract_wave_data(texts):
    """WAVE情報を抽出してすべての該当情報を出力する"""
    wave_pattern = re.compile(r'(WAVE \d|EX-WAVE)')  # WAVE 1, WAVE 2, WAVE 3, EX-WAVEをすべてマッチ
    result_pattern = re.compile(r'(GJ!|NG)')
    gold_eggs_pattern = re.compile(r'\b([1-9][0-9]{0,2})/([1-9][0-9]?)\b')
    tide_pattern = re.compile(r'(満潮|普通|干潮)')

    wave_names = []
    wave_results = []
    gold_eggs_set = set()  # 重複を避けるためにsetを使用
    tides = []

    for text in texts:
        description = text.description.strip()

        # まず日付形式の部分を除外
        description = re.sub(r'\d{4}/\d{1,2}/\d{1,2}', '', description)

        # WAVE名の抽出
        wave_match = wave_pattern.findall(description)  # 全てのWAVEパターンを抽出
        if wave_match:
            wave_names.extend(wave_match)  # 複数マッチした場合もリストに追加

        # WAVE結果
        result_match = result_pattern.findall(description)
        if result_match:
            wave_results.extend(result_match)

        # 金イクラの数をフィルタリングして抽出
        gold_eggs_match = gold_eggs_pattern.findall(description)
        if gold_eggs_match:
            for gold_egg in gold_eggs_match:
                gold_eggs_set.add(f"{gold_egg[0]}/{gold_egg[1]}")  # 抽出された値を追加

        # 潮の状況
        tide_match = tide_pattern.findall(description)
        if tide_match:
            tides.extend(tide_match)

    # ノルマの数に基づいて並べ替え（昇順）
    sorted_gold_eggs = sorted(list(gold_eggs_set), key=lambda x: int(x.split('/')[1]), reverse=False)

    # 潮の数をWAVE名の数に合わせる
    if len(tides) > len(wave_names):
        tides = tides[:len(wave_names)]  # 潮が多すぎる場合はカット
    elif len(tides) < len(wave_names):
        tides.extend(['?'] * (len(wave_names) - len(tides)))  # 足りない場合は'?'で埋める

    return {
        "wave_names": wave_names,
        "wave_results": wave_results,
        "gold_eggs": sorted_gold_eggs,  # ソートされたリストを返す
        "tides": tides
    }

# シナリオコードを抽出する関数
def extract_scenario_code(texts):
    """シナリオコードを抽出する"""
    scenario_code_pattern = re.compile(r'S[A-Z0-9]{15}')  # シナリオコードのパターン
    for text in texts:
        description = text.description.strip()
        match = scenario_code_pattern.search(description)
        if match:
            return match.group()
    return "シナリオコードが見つかりません"

# Googleスプレッドシートにデータを転記する関数
def write_to_google_sheet(data, sheet_url):
    try:
        # Google Sheets APIの認証情報
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        creds = Credentials.from_service_account_file(json_file_path, scopes=SCOPES)

        # Google Sheets APIに接続
        client = gspread.authorize(creds)

        # スプレッドシートの取得（URLからスプレッドシートを開く）
        sheet = client.open_by_url(sheet_url).worksheet('読み取り結果')

        # データをGoogle Sheetsに追加
        sheet.append_row([
            data['message_timestamp'],   # 送信日時
            data['discord_user_id'],     # Discordユーザ名
            data['thread_name'],         # スレッド名
            data['image_url'],           # 画像URL
            ', '.join(data['date_time']), # 日付と時間
            data['stage_name'],          # ステージ名
            ', '.join(data['clear_status']),  # クリア状況
            ', '.join(data['danger_rate']),   # キケン度
            ', '.join(data['wave_names']),    # WAVE名
            ', '.join(data['gold_eggs']),     # 金イクラの数
            ', '.join(data['tides']),         # 潮
            data['scenario_code'],            # シナリオコード
            data['step_value'],               # すてっぷ数（半角数字）
            data['judge']                     # 受理・不受理結果
        ])

    except gspread.exceptions.APIError as e:
        # API エラーの詳細を出力
        print(f"Google Sheets API エラー: {e}")
        print(f"レスポンス: {e.response.text}")
        
    except requests.exceptions.JSONDecodeError as e:
        # JSON デコードエラーの詳細を出力
        print(f"JSON デコードエラー: {e.msg}")
        print(f"原因となったレスポンス: {e.doc}")
        
    except Exception as e:
        # その他のエラーの詳細を出力
        print(f"その他のエラー: {e}")


# チームリストをGoogleスプレッドシートから取得する関数
def get_team_list(sheet_url):
    # Google Sheets APIの認証情報
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    creds = Credentials.from_service_account_file(json_file_path, scopes=SCOPES)

    # Google Sheets APIに接続
    client = gspread.authorize(creds)

    # チーム名リストシートの取得
    sheet = client.open_by_url(sheet_url).worksheet('チーム名リスト')

    # チーム名リストのデータを取得
    data = sheet.get_all_records()

    # データを確認（デバッグ用）
    #print(data)  # スプレッドシートの内容をターミナルに出力して確認

    return data


# シナリオコードを抽出する関数
def extract_scenario_code(texts):
    """シナリオコードを抽出する"""
    scenario_code_pattern = re.compile(r'S[A-Z0-9]{15}')  # シナリオコードのパターン
    for text in texts:
        description = text.description.strip()
        match = scenario_code_pattern.search(description)
        if match:
            return match.group()
    return "シナリオコードが見つかりません"

# キケン度に基づいて「すてっぷ」を計算する関数
def calculate_step(danger_rate_str):
    """キケン度文字列を元にすてっぷを計算する"""
    print(f"Received danger_rate_str: {danger_rate_str}")  # デバッグ用のログを追加

    # 「キケン」または「キケン度」の後に1～3桁の数字が続くパターンを正規表現でキャッチ
    match = re.search(r'キケン[度\s]*([0-9]{1,3})%', danger_rate_str)  # 「キケン」＋数字を抽出
    if match:
        try:
            danger_rate = int(match.group())  # 抽出された数値をintに変換
            print(f"Extracted danger_rate: {danger_rate}")  # デバッグ用ログ

            # ステップ数の判定ロジック
            if 0 <= danger_rate <= 60:
                return 1
            elif 61 <= danger_rate <= 120:
                return 2
            elif 121 <= danger_rate <= 150:
                return 3
            elif 151 <= danger_rate <= 180:
                return 4
            elif 181 <= danger_rate <= 200:
                return 5
            elif 201 <= danger_rate <= 220:
                return 6
            elif 221 <= danger_rate <= 250:
                return 7
            elif 251 <= danger_rate <= 270:
                return 8
            elif 271 <= danger_rate <= 300:
                return 9
            elif 301 <= danger_rate <= 333:
                return 10
            else:
                print(f"Unexpected danger_rate value: {danger_rate}")  # 範囲外の数値があった場合
        except ValueError:
            print("Failed to convert danger_rate to int.")  # 数値変換に失敗した場合
    else:
        print("Failed to extract danger_rate.")  # 正規表現にマッチしなかった場合

    return None  # キケン度が取得できなかった場合や範囲外の場合

def determine_step_value(scenario_code, danger_rate_str):
    """シナリオコードに基づいてステップ数を決定し、該当しない場合はキケン度に基づいて判定"""
    
    # シナリオコードごとのステップ数マッピング
    scenario_code_to_step = {
        "": 1,#適切に指定してください
        "": 2,
        "": 3,
        "": 4,
        "": 5,
        "": 6,
        "": 7,
        "": 8,
        "": 9,
        "": 10
    }

    # シナリオコードが一致すればそのステップ数を返す
    if scenario_code in scenario_code_to_step:
        return scenario_code_to_step[scenario_code]
    
    # シナリオコードが一致しない場合は、キケン度に基づいてステップ数を決定
    return calculate_step(danger_rate_str)

# メッセージが送信されたときのイベント
@bot.event
async def on_message(message):
    if message.author == bot.user:
        return

    # メッセージがスレッド内で送信されたかを確認
    if isinstance(message.channel, discord.Thread):
        thread = message.channel
        thread_name = thread.name  # スレッド名を取得
    else:
        # スレッド以外（通常のチャンネル）で画像が送信された場合はエラーメッセージを返す
        if message.attachments:
            await message.reply("用意された専用スレッドから画像を送付してください。")
        return  # スレッド以外では解析を行わない

    # メッセージに画像が含まれているか確認
    if message.attachments:
        image_data_list = []
        date_time = None  # 1枚目の画像で日付を取得
        stage_name = None  # 1枚目のステージ名を保持
        clear_status = None  # 1枚目のクリア状況を保持
        danger_rate = None  # 1枚目のキケン度を保持
        wave_data = None  # 1枚目のWAVE情報を保持
        scenario_code = None  # シナリオコードを保持
        judge = "不受理"  # 初期値は「不受理」

# キケン度に基づいて「すてっぷ」を計算する関数
def calculate_step(danger_rate_str):
    """キケン度文字列を元にすてっぷを計算する"""
    print(f"Received danger_rate_str: {danger_rate_str}")  # デバッグ用のログを追加

    # キケン度に関する数値を抽出する正規表現（「キケン」と1～3桁の数字を取得）
    match = re.search(r'\b\d{1,3}\b', danger_rate_str)
    if match:
        danger_rate = int(match.group())  # 抽出された整数部分を変数に格納
        print(f"Extracted danger_rate: {danger_rate}")  # デバッグ用ログ

        # ステップ数を判定するロジック
        if 0 <= danger_rate <= 60:
            return 1
        elif 61 <= danger_rate <= 90:
            return 2
        elif 91 <= danger_rate <= 120:
            return 3
        elif 121 <= danger_rate <= 150:
            return 4
        elif 151 <= danger_rate <= 200:
            return 5
        elif 201 <= danger_rate <= 224:
            return 6
        elif 225 <= danger_rate <= 250:
            return 7
        elif 251 <= danger_rate <= 286:
            return 8
        elif 287 <= danger_rate <= 301:
            return 9
        elif 302 <= danger_rate <= 333:
            return 10
    else:
        print("Failed to extract danger_rate.")  # デバッグ用ログ

    return None  # キケン度が取得できなかった場合や範囲外の場合



# チーム名リストの更新関数
def update_team_list(sheet_url, team_name, step_value, judge_result):
    """チーム名リストのすてっぷ列に judge 結果を更新"""
    # Google Sheets APIの認証情報
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    creds = Credentials.from_service_account_file(json_file_path, scopes=SCOPES)
    client = gspread.authorize(creds)

    # チーム名リストシートの取得
    sheet = client.open_by_url(sheet_url).worksheet('チーム名リスト')
    
    # シート全体のデータを取得
    data = sheet.get_all_values()
    
    # チーム名を検索して行番号を特定
    row_number = None
    for i, row in enumerate(data):
        if team_name in row[1]:  # 2列目がチーム名列と仮定
            row_number = i + 1  # Google Sheetsの行番号は1から始まる
            break
    
    if row_number:
        # step_value が None の場合、処理をスキップ
        if step_value is None:
            print(f"step_value is None for team {team_name}. Skipping update.")
            return False  # 更新が行われなかったことを示す
        
        # すてっぷに対応する列（4列目以降がすてっぷ列）
        step_column = step_value + 3  # すてっぷ1は4列目なので+3
        sheet.update_cell(row_number, step_column, judge_result)  # 該当セルを更新
        return True
    return False


# チームの画像受理結果を取得する関数
def get_team_status(sheet_url, team_name):
    """チームの画像受理結果を取得し、受理状況を返す"""
    # Google Sheets APIの認証情報
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    creds = Credentials.from_service_account_file(json_file_path, scopes=SCOPES)
    client = gspread.authorize(creds)

    # チーム名リストシートの取得
    sheet = client.open_by_url(sheet_url).worksheet('チーム名リスト')

    # シート全体のデータを取得
    data = sheet.get_all_values()

    # チーム名を検索して行番号を特定
    row_number = None
    for i, row in enumerate(data):
        if row[1] == team_name:  # 2列目がチーム名列
            row_number = i + 1  # Google Sheetsの行番号は1から始まる
            break

    if row_number:
        # 該当するチームの受理結果（4列目以降がすてっぷ列）
        step_status = data[row_number - 1][3:13]  # すてっぷ1は4列目（インデックス3）、すてっぷ10まで取得

        # 受理状況をリスト形式で返す（空白も含めて）
        status_list = [f"すてっぷ{index + 1}: {status or '未提出'}" for index, status in enumerate(step_status)]
        return status_list, step_status  # ステータスリストと元のステータスを返す
    return [], []

@bot.event
async def on_message(message):
    if message.author == bot.user:
        return

    # メッセージがスレッド内で送信されたかを確認
    if isinstance(message.channel, discord.Thread):
        thread = message.channel
        thread_name = thread.name  # スレッド名を取得
        parent_channel = thread.parent  # スレッドの親チャンネルを取得
        
        # /yomitorikunで有効化されたチャンネルでなければ処理をスキップ
        if parent_channel.id not in active_channels:
            return

    else:
        # スレッド以外のチャンネルで、/yomitorikunコマンドで有効化されたチャンネル以外なら何もしない
        if message.channel.id not in active_channels:
            return
        
        # 画像が送信された場合
        if message.attachments:
            await message.reply("用意された専用スレッドから画像を送付してください。")
            await message.delete()  # メッセージを削除する処理を追加
        return  # スレッド以外では解析を行わない

    # メッセージに画像が含まれているか確認
    if message.attachments:
        image_data_list = []
        date_time = None  # 1枚目の画像で日付を取得
        stage_name = None  # 1枚目のステージ名を保持
        clear_status = None  # 1枚目のクリア状況を保持
        danger_rate = None  # 1枚目のキケン度を保持
        wave_data = None  # 1枚目のWAVE情報を保持
        scenario_code = None  # シナリオコードを保持
        step_value = None  # すてっぷを保持
        judge = "不受理"  # 初期値として不受理を設定


        # Discordの送信日時（UTC）をJSTに変換
        utc_timestamp = message.created_at.replace(tzinfo=pytz.utc)  # UTCのタイムゾーンを設定
        jst_timestamp = utc_timestamp.astimezone(jst)  # JSTに変換
        message_timestamp = jst_timestamp.strftime("%Y-%m-%d %H:%M:%S")  # JST形式でフォーマット

        image_urls = []  # 画像URLを保存するリスト

        for idx, attachment in enumerate(message.attachments):
            if attachment.content_type.startswith('image/'):
                image_url = attachment.url
                image_urls.append(image_url)  # 画像URLをリストに追加
                response = requests.get(image_url)
                image_path = f'temp_image_{idx}.png'  # 複数画像に対応

                with open(image_path, 'wb') as file:
                    file.write(response.content)

                # 画像をGoogle Vision APIで処理してテキスト情報を抽出
                texts, all_texts = detect_text(image_path)

                # 日付と時間、クリア状況、キケン度のパターン
                date_time_pattern = re.compile(r'\d{4}/\d{1,2}/\d{1,2} \d{2}:\d{2}')
                clear_status_pattern = re.compile(r'(clear!{0,3}|failure)', re.IGNORECASE)
                danger_rate_pattern = re.compile(r'キケン(度)? \d{2,3}%')

                if idx == 0:
                    # 1枚目の画像をチェックして日付を確認
                    date_time = extract_specific_info(texts, date_time_pattern)
                    if not date_time:
                        # 日付が読み取れなかった場合のエラーメッセージをリプライ形式で返して処理を終了
                        await message.reply("日付が読み取れませんでした。クリア日時を含むリザルト画像を1枚目にして送付してください。")
                        return
                    else:
                        # 1枚目で日付が読み取れた場合
                        clear_status = extract_specific_info(texts, clear_status_pattern)

                        # クリア状況の判定処理を追加
                        if any(re.match(r'clear!{0,3}', cs, re.IGNORECASE) for cs in clear_status):
                            clear_status = ["Clear!!"]  # clear系の表現をClear!!に置き換え
                        elif any(re.match(r'failure', cs, re.IGNORECASE) for cs in clear_status):
                            clear_status = ["Failure"]  # failure系の表現をFailureに置き換え

                        danger_rate = extract_specific_info(texts, danger_rate_pattern)
                        stage_name = extract_stage_name(texts)
                        wave_data = extract_wave_data(texts)
                        scenario_code = extract_scenario_code(texts)  # シナリオコードを抽出

                        # ステップ数の判定
                        if scenario_code:
                            step_value = determine_step_value(scenario_code, danger_rate[0] if danger_rate else '')
                        else:
                            step_value = calculate_step(danger_rate[0] if danger_rate else '')

                        # 開始日時と終了日時が設定されていれば、日付を範囲内かチェック
                        if start_dt and end_dt:
                            image_dt = datetime.strptime(date_time[0], '%Y/%m/%d %H:%M')
                            if start_dt <= image_dt <= end_dt and 'Clear!!' in clear_status:
                                judge = "受理"  # 条件が合えば受理
                        elif 'Clear!!' in clear_status:
                            judge = "受理"  # クリア状況がOKであれば受理

                        # 解析結果を元のメッセージへの返信として送信
                        result_message = (
                            #f"抽出されたテキスト:\n{all_texts}\n\n"
                            f"日付と時間: {', '.join(date_time)}\n"
                            #f"ステージ名: {stage_name}\n"
                            f"クリア状況: {', '.join(clear_status)}\n"
                            f"キケン度: {', '.join(danger_rate)}\n"
                            f"すてっぷ: {step_value}\n"
                            f"WAVE名: {', '.join(wave_data.get('wave_names', []))}\n"
                            f"金イクラの数: {', '.join(wave_data.get('gold_eggs', []))}\n"
                            f"シナリオコード: {scenario_code}\n"
                            f"画像の判定結果: {judge}"
                        )

                        await message.reply(result_message)

                        # 1枚目のデータを保存
                        image_data_list.append({
                            'thread_name': thread_name,  # スレッド名を保存
                            'date_time': date_time,
                            'stage_name': stage_name,
                            'clear_status': clear_status,
                            'danger_rate': danger_rate,
                            'wave_results': wave_data.get('wave_results', []),
                            'wave_names': wave_data.get('wave_names', []),
                            'gold_eggs': wave_data.get('gold_eggs', []),
                            'tides': wave_data.get('tides', []),
                            'scenario_code': scenario_code,  # シナリオコードを追加
                            'discord_user_id': message.author.display_name,
                            'image_url': image_url,  # 1枚目の画像URLを保存
                            'message_timestamp': message_timestamp,
                            'step_value': step_value,  # すてっぷを保存
                            'judge': judge  # 受理・不受理を保存
                        })
                else:
                    # 2枚目以降の画像はURLだけ保存し、それ以外の情報は1枚目のものを使用
                    image_data_list[0]['image_url'] += f", {image_url}"  # URLを追記
                    # 2枚目以降は、シートに記録するための情報は追記しない

        # Googleスプレッドシートに転記
        if sheet_url:
            for game_data in image_data_list:
                write_to_google_sheet(game_data, sheet_url)  # 最初のスプレッドシート転記
        else:
            await message.channel.send("スプレッドシートのURLが設定されていません。/sheet コマンドを使用して設定してください。")

        # チーム名を取得（スレッド名の形式が No_チーム名_タイムスタンプ）し、受理状況を更新
        team_name = thread_name.split('_')[1]

        # Googleスプレッドシートに受理結果を追記
        if update_team_list(sheet_url, team_name, step_value, judge):
            # 更新が成功したらチームの画像受理状況を取得
            team_status, raw_status = get_team_status(sheet_url, team_name)

            # すべてのステップが「受理」かを確認
            if all(status == "受理" for status in raw_status[:10]):
                success_message = "\n".join(team_status) + "\n\nすべての画像が提出されました。\n" \
                                 "なお、本システムは簡易な解析のみを行なっており誤検出もあり得ます。\n" \
                                 "提出画像に不備（枚数、クリア状況など）がないかなど、改めて自己責任でご確認ください。"
                await message.reply(success_message)
            else:
                status_message = "現在の画像受理状況です。\n" + "\n".join(team_status)
                await message.reply(status_message)
        else:
            await message.reply(f"{team_name} の情報を更新できませんでした。")


# Discordボットの実行
bot.run(TOKEN)
