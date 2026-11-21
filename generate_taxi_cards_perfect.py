import os
from PIL import Image, ImageFont, ImageDraw

# 字型路徑
FONT_JP_BOLD = "C:/Windows/Fonts/meiryob.ttc"
FONT_JP_REG = "C:/Windows/Fonts/meiryo.ttc"
FONT_ZH_BOLD = "C:/Windows/Fonts/msjhbd.ttc"
FONT_ZH_REG = "C:/Windows/Fonts/msjh.ttc"

def get_font(path, pt):
    # pt 轉 px (96 DPI)
    return ImageFont.truetype(path, int(pt * 96 / 72))

def get_fit_font(font_path, text, target_pt, max_width, min_pt=15):
    pt = target_pt
    while pt >= min_pt:
        font = get_font(font_path, pt)
        bbox = font.getbbox(text)
        if (bbox[2] - bbox[0]) <= max_width:
            return font, pt
        pt -= 1
    return get_font(font_path, min_pt), min_pt

def draw_taxi_card(data, output_path):
    width = 1000
    height = 1400

    img = Image.new("RGBA", (width, height), (248, 250, 252, 255))
    draw = ImageDraw.Draw(img)

    # 1. Top Header Banner
    draw.rectangle([0, 0, width, 140], fill=(15, 35, 75, 255))

    # Header Tag
    tag_font = get_font(FONT_JP_BOLD, 18)
    draw.text((40, 25), data['tag'], font=tag_font, fill=(254, 215, 0, 255))

    # Header Title (Fit to width 920)
    title_font, _ = get_fit_font(FONT_JP_BOLD, data['title'], 28, 920, min_pt=20)
    draw.text((40, 68), data['title'], font=title_font, fill=(255, 255, 255, 255))

    # 2. Greeting / Request Box (White Card)
    # y: 160 ~ 315 (height 155)
    draw.rounded_rectangle([40, 160, 960, 315], radius=12, fill=(255, 255, 255, 255), outline=(210, 220, 235, 255), width=2)
    req_jp_font, _ = get_fit_font(FONT_JP_BOLD, data['request_jp'], 23, 880, min_pt=17)
    draw.text((60, 185), data['request_jp'], font=req_jp_font, fill=(15, 23, 42, 255))

    req_zh_font, _ = get_fit_font(FONT_ZH_REG, data['request_zh'], 17, 880, min_pt=14)
    draw.text((60, 255), data['request_zh'], font=req_zh_font, fill=(100, 116, 139, 255))

    # 3. SECTION 1: 回程時間 (Time) Highlight Box
    # y: 335 ~ 540 (height 205)
    draw.rounded_rectangle([40, 335, 960, 540], radius=12, fill=(255, 243, 205, 255), outline=(245, 158, 11, 255), width=3)
    
    sec_label_font = get_font(FONT_JP_BOLD, 16)
    draw.text((60, 350), "⏰ お迎え希望時間（回程接送時間）", font=sec_label_font, fill=(180, 83, 9, 255))

    time_val_font, _ = get_fit_font(FONT_JP_BOLD, data['time_text'], 40, 880, min_pt=26)
    draw.text((60, 395), data['time_text'], font=time_val_font, fill=(185, 28, 28, 255))

    time_sub_font, _ = get_fit_font(FONT_ZH_BOLD, data['time_sub'], 18, 880, min_pt=14)
    draw.text((60, 480), data['time_sub'], font=time_sub_font, fill=(71, 85, 105, 255))

    # 4. SECTION 2: 乘車地點 (Location) Highlight Box
    # y: 560 ~ 765 (height 205)
    draw.rounded_rectangle([40, 560, 960, 765], radius=12, fill=(232, 244, 253, 255), outline=(2, 132, 199, 255), width=3)

    draw.text((60, 575), "📍 乗車場所・待ち合わせ場所（上車候車地點）", font=sec_label_font, fill=(3, 105, 161, 255))

    loc_val_font, _ = get_fit_font(FONT_JP_BOLD, data['loc_text'], 27, 880, min_pt=19)
    draw.text((60, 620), data['loc_text'], font=loc_val_font, fill=(15, 23, 42, 255))

    loc_sub_font, _ = get_fit_font(FONT_ZH_BOLD, data['loc_sub'], 18, 880, min_pt=14)
    draw.text((60, 705), data['loc_sub'], font=loc_sub_font, fill=(71, 85, 105, 255))

    # 5. SECTION 3: 詢問司機是否在此等候 (Question)
    # y: 785 ~ 935 (height 150)
    draw.rounded_rectangle([40, 785, 960, 935], radius=12, fill=(255, 255, 255, 255), outline=(210, 220, 235, 255), width=2)

    q_font, _ = get_fit_font(FONT_JP_BOLD, data['question_jp'], 23, 880, min_pt=17)
    draw.text((60, 810), data['question_jp'], font=q_font, fill=(30, 41, 59, 255))

    q_zh_font, _ = get_fit_font(FONT_ZH_REG, data['question_zh'], 17, 880, min_pt=14)
    draw.text((60, 878), data['question_zh'], font=q_zh_font, fill=(100, 116, 139, 255))

    # 6. SECTION 4: 司機手指回覆區 (Driver Response)
    # y: 955 ~ 1335 (height 380)
    draw.rounded_rectangle([40, 955, 960, 1335], radius=12, fill=(240, 253, 244, 255), outline=(34, 197, 94, 255), width=3)

    draw.text((60, 970), "👉 運転手さんへ：こちらを指差してください（請司機先生手指回覆）", font=sec_label_font, fill=(21, 128, 61, 255))

    opt_y = 1010
    box_height = 92
    box_gap = 14

    for i in range(len(data['options_jp'])):
        # Option Box (White)
        draw.rounded_rectangle([60, opt_y, 940, opt_y + box_height], radius=8, fill=(255, 255, 255, 255), outline=(187, 247, 208, 255), width=2)

        # Checkmark Box
        draw.rounded_rectangle([80, opt_y + 24, 122, opt_y + 66], radius=6, fill=(240, 253, 244, 255), outline=(34, 197, 94, 255), width=3)

        # Option JP Text (max width 780: 920 - 140 = 780)
        opt_jp_font, _ = get_fit_font(FONT_JP_BOLD, data['options_jp'][i], 20, 780, min_pt=15)
        draw.text((140, opt_y + 14), data['options_jp'][i], font=opt_jp_font, fill=(15, 23, 42, 255))

        # Option ZH Text
        opt_zh_font, _ = get_fit_font(FONT_ZH_REG, data['options_zh'][i], 15, 780, min_pt=12)
        draw.text((140, opt_y + 54), data['options_zh'][i], font=opt_zh_font, fill=(100, 116, 139, 255))

        opt_y += box_height + box_gap

    # 7. Footer Note
    foot_font = get_font(FONT_ZH_REG, 14)
    draw.text((40, 1355), "※ 提示：請於上車或抵達目的地時，將此畫面出示給計程車司機確認。", font=foot_font, fill=(148, 163, 184, 255))

    # Save PNG
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    img.save(output_path, "PNG")
    print(f"Generated successfully: {output_path}")

cards = [
    # 1. Day 2 鬼之城
    {
        'output': 'timetables/taxi_return_kinojo.png',
        'tag': 'Day 2 【鬼ノ城 ➔ 総社駅】',
        'title': '帰りのタクシー予約・乗車場所の確認',
        'request_jp': '運転手さんへ：帰りのタクシー（お迎え）をお願いできますか？',
        'request_zh': '司機先生您好：請問可以向您預約回程下山的計程車嗎？',
        'time_text': '10 : 30  お迎え希望',
        'time_sub': '（午前 10:30 に総社駅へ戻りたいです）',
        'loc_text': '鬼城山ビジターセンター 駐車場（ここ）',
        'loc_sub': '（上車處：就在鬼城山遊客中心前的停車場）',
        'question_jp': '10:30 に この場所（駐車場）でお待ちすればよろしいですか？',
        'question_zh': '我們 10:30 在這個停車場等您來接，可以嗎？',
        'options_jp': [
            'はい、10:30にここへ迎えに来ます（予約了解）',
            '別のタクシーを手配します（後ほど来ます）',
            '予約できません（下山時に電話で呼んでください）'
        ],
        'options_zh': [
            '好的，我會在 10:30 到這裡接您（預約確認）',
            '我會幫您派另一輛車前來接您',
            '目前無法預約，下山時請致電叫車'
        ]
    },
    # 2. Day 8 元乃隅神社
    {
        'output': 'timetables/taxi_return_motonosumi.png',
        'tag': 'Day 8 【元乃隅神社 ➔ 長門古市駅】',
        'title': '帰りのタクシー予約・乗車場所の確認',
        'request_jp': '運転手さんへ：長門古市駅への帰りのタクシーをお願いできますか？',
        'request_zh': '司機先生您好：請問能預約回程接我們回長門古市站嗎？',
        'time_text': '11 : 35  お迎え希望（必須）',
        'time_sub': '（※ 12:09 発の JR 列車に接続するため、11:35 出発が必須です！）',
        'loc_text': '元乃隅神社 第1駐車場（鳥居入口・ここ）',
        'loc_sub': '（上車處：元乃隅神社 第1停車場 / 鳥居入口前）',
        'question_jp': '11:35 に この場所（第1駐車場）でお待ちすればよろしいですか？',
        'question_zh': '我們 11:35 在第1停車場等您來接，可以嗎？',
        'options_jp': [
            'はい、11:35に迎えに来ます（12:09列車に接続OK）',
            '別の車を手配します（11:35に同僚が来ます）',
            '予約できません（電話連絡が必要です）'
        ],
        'options_zh': [
            '好的，我會在 11:35 來接您（可趕上 12:09 火車）',
            '我會安排另一位同事於 11:35 來接您',
            '無法預約，回程時需電話聯絡叫車'
        ]
    },
    # 3. Day 13 備中松山城 乘合計程車
    {
        'output': 'timetables/taxi_return_bitchu.png',
        'tag': 'Day 13 【備中松山城 ふいご峠 ➔ 備中高梁駅】',
        'title': '観光乗合タクシー・お迎え確認',
        'request_jp': '乗合タクシー運転手さんへ：帰りの便の乗車確認をお願いします。',
        'request_zh': '乘合計程車司機先生您好：想確認回程搭車時間與地點。',
        'time_text': '13 : 10  発（ふいご峠発）',
        'time_sub': '（※ 備中高梁駅 13:45 発の JR 列車に接続します）',
        'loc_text': 'ふいご峠 駐車場・乗合タクシー乗降場（ここ）',
        'loc_sub': '（上車處：八合目鞴峠登山口 停車場乘車處）',
        'question_jp': '13:10 に この場所（ふいご峠）でお待ちすれば乗車できますか？',
        'question_zh': '我們 13:10 在鞴峠此處候車，即可順利搭車嗎？',
        'options_jp': [
            'はい、13:10にここでお待ちください（乗車OK）',
            '予約のお名前を確認させてください',
            '別の便（13:40発など）に変更しますか？'
        ],
        'options_zh': [
            '好的，請於 13:10 在此處等候上車（乘車確認）',
            '請向我確認您的預約大名',
            '是否需要改搭其他班次（例如 13:40 發車）？'
        ]
    },
    # 4. 通用版
    {
        'output': 'timetables/taxi_return_general.png',
        'tag': '【タクシー往復・指さし確認用】',
        'title': 'タクシーお迎え予約・乗車場所の確認',
        'request_jp': '運転手さんへ：後ほど、帰りの迎えに来ていただくことは可能ですか？',
        'request_zh': '司機先生您好：請問稍後能請您開車來接我們回程嗎？',
        'time_text': 'お迎え希望時間： [　　：　　]',
        'time_sub': '（指で時間を指すか、スマホの時計をお見せします）',
        'loc_text': 'ここ（降りる場所 / 駐車場）でお待ちします',
        'loc_sub': '（若有指定其他候車處，請司機先生直接指點）',
        'question_jp': '帰りの乗車場所（待ち合わせ場所）はここになりますか？',
        'question_zh': '請問回程在此處（或指定地點）等您來接，可以嗎？',
        'options_jp': [
            'はい、指定の時間にここへ迎えに来ます（予約了解）',
            'ここで待っていてください（同じ場所で乗車）',
            'すみません、予約できません（電話で呼んでください）'
        ],
        'options_zh': [
            '好的，會在指定時間來這裡接您（預約確認）',
            '請在現在下車的這個地方等候即可',
            '抱歉目前無法預約，請回程時打電話叫車'
        ]
    }
]

for card in cards:
    draw_taxi_card(card, card['output'])
