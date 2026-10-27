
def format_word_date_with_dow(date_str):
    if not date_str:
        return ''
    try:
        from datetime import datetime
        dt = datetime.strptime(date_str.split(' ')[0], '%Y-%m-%d')
        weekdays = ['一', '二', '三', '四', '五', '六', '日']
        return f"{date_str} ({weekdays[dt.weekday()]})"
    except Exception:
        return date_str

import re
# -*- coding: utf-8 -*-
"""
山陰山陽 13 天旅行小冊子 Word (.docx) 產生器
每日前後自成獨立篇章，具備強制分頁、景點詳細資訊、精準車次時刻、車資預算表及紀念章蓋印框。
"""

import os
import json
import io
import base64
from PIL import Image
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def get_spot_display_name(sid, spots_map, default_val=None):
    if not sid:
        return default_val or ""
    if sid in spots_map:
        return spots_map[sid].get('name', sid)
    name_dict = {
        'okayama_airport': '岡山桃太郎機場 (OKJ)',
        'okayama_hotel': 'VIA INN 岡山',
        'tottori_hotel': 'APA飯店 鳥取車站前',
        'kinosaki_onsen': '城崎溫泉 和風旅館',
        'ine_funaya': '伊根舟屋聚落',
        'yonago_hotel': 'Hotel Harvest In Yonago',
        'yasugi_lodging': '安來站前精選飯店',
        'kaike_onsen': '皆生溫泉 海濱旅館',
        'tamatsukuri_onsen': '玉造溫泉 和風旅館',
        'hagi_lodging': '萩市區飯店 / 萩溫泉旅館',
        'shin_yamaguchi_lodging': 'Comfort Hotel Shin Yamaguchi',
        'hiroshima_lodging_d8': 'Via Inn Hiroshima Shinkansenguchi',
        'hiroshima_lodging_d9': 'Via Inn Hiroshima Shinkansenguchi',
        'tomonoura_onsen': '鞆之浦溫泉 遠景之宿',
        'daishoin': '宮島彌山 大本山 大聖院',
    }
    return name_dict.get(sid, default_val or sid)

def get_segment_display_title(key, day, spots_map):
    d_num = day.get('dayNumber')
    if key == f'tpe_flight_{d_num}' or 'tpe' in key:
        return '台灣桃園機場 ➔ 啟程搭機出發'
    if key == f'okayama_airport_{d_num}':
        return '抵達岡山桃太郎機場 (IT214 航班)'
    if key == f'morning_origin_{d_num}':
        return f"Day {d_num} 晨間出發地 ({day.get('baseStation', '').split(' / ')[0]})"
    
    if '_to_' in key:
        parts = key.split('_to_')
        from_sid, to_sid = parts[0], parts[1]
        from_name = get_spot_display_name(from_sid, spots_map, from_sid)
        to_name = get_spot_display_name(to_sid, spots_map, to_sid)
        return f"{from_name} ➔ {to_name}"
        
    return f"搭乘區間 ({key})"

def insert_timetable_image(doc, img_url, img_name="時刻表縮圖", max_width_in=5.8, max_height_in=4.6):
    """將時刻表圖檔（支援檔案路徑與 base64 data-url）自適應加入 Word，不超出版面"""
    stream = None
    is_stream_obj = False
    if img_url.startswith('data:'):
        try:
            _, b64_str = img_url.split(',', 1)
            img_bytes = base64.b64decode(b64_str)
            stream = io.BytesIO(img_bytes)
            is_stream_obj = True
        except Exception:
            return False
    elif os.path.exists(img_url):
        stream = img_url
    else:
        return False

    try:
        # 計算等比例縮放避免超出頁面
        calc_w = Inches(max_width_in)
        try:
            with Image.open(stream) as im:
                orig_w, orig_h = im.size
                if orig_w > 0 and orig_h > 0:
                    aspect = orig_h / orig_w
                    if aspect > (max_height_in / max_width_in):
                        # 直式長圖：限制高度
                        calc_w = Inches(max_height_in / aspect)
                    else:
                        calc_w = Inches(max_width_in)
        except Exception:
            pass

        if is_stream_obj and hasattr(stream, 'seek'):
            stream.seek(0)

        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run()
        run.add_picture(stream, width=calc_w)

        # 圖片檔名標註
        p_caption = doc.add_paragraph()
        p_caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_caption.paragraph_format.space_before = Pt(0)
        p_caption.paragraph_format.space_after = Pt(4)
        r_cap = p_caption.add_run(f"📎 附件圖檔：{img_name}")
        format_run(r_cap, size_pt=7.5, color_rgb=(100, 116, 139))
        return True
    except Exception as e:
        print(f"插入圖檔警告: {img_url}, 原因: {e}")
        return False


def set_cell_background(cell, fill_hex):
    """設定儲存格背景底色"""
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """設定儲存格內部 padding (單位: dxa, 20 dxa = 1 pt)"""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_cell_border(cell, **kwargs):
    """
    設定儲存格框線
    kwargs: top, bottom, left, right 各自為 dict(sz=12, val='single', color='CBD5E1')
    """
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for border_name, border_props in kwargs.items():
        node = OxmlElement(f'w:{border_name}')
        for key, val in border_props.items():
            node.set(qn(f'w:{key}'), str(val))
        tcBorders.append(node)
    tcPr.append(tcBorders)

def add_header_footer(doc):
    """設定全份文件頁首與頁尾"""
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.9)
        section.right_margin = Inches(0.9)
        section.different_first_page_header_footer = True
        
        # 頁首
        header = section.header
        hp = header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hrun = hp.add_run("山陰山陽壯遊 13 天・旅行隨身手冊")
        hrun.font.name = "Microsoft JhengHei"
        hrun.font.size = Pt(8.5)
        hrun.font.color.rgb = RGBColor(148, 163, 184)
        
        # 頁尾
        footer = section.footer
        fp = footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        frun = fp.add_run("— 旅行隨身手冊 —")
        frun.font.name = "Microsoft JhengHei"
        frun.font.size = Pt(8.5)
        frun.font.color.rgb = RGBColor(148, 163, 184)

def format_run(run, font_name="Microsoft JhengHei", size_pt=10.5, bold=False, color_rgb=(51, 65, 85), italic=False):
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = RGBColor(*color_rgb)

def get_short_day_title(day, spots_map):
    d_num = day.get('dayNumber', 1)
    spot_ids = day.get('spotIds', [])
    names = []
    
    custom_names = {
        'taoyuan_airport': '桃園機場',
        'okayama_airport': '岡山桃太郎機場',
        'okonomiyaki_kamons': '岡山御好燒花門S',
        'aeon_mall_okayama': 'AEON Mall',
        'okayama_hotel': '岡山',
        'kinojo': '鬼ノ城',
        'chiyagyu_bisai': '鉄板焼 備彩',
        'tsuyama': '津山城',
        'tottori_hotel': '鳥取',
        'tottori': '鳥取城',
        'tottori_sakyu': '鳥取砂丘',
        'tottori_takumi': 'たくみ割烹店',
        'kinosaki_onsen': '城崎溫泉',
        'amanohashidate_boat': '天橋立遊覽船',
        'amanohashidate': '天橋立',
        'hashidate_chaya': 'はしだて茶屋',
        'ine_funaya': '伊根舟屋',
        'araiso_yonago': '荒磯 本店',
        'yonago_hotel': '米子',
        'tsuwano': '津和野城',
        'taikodani_inari': '太鼓谷稻成神社',
        'kaike_onsen': '皆生溫泉',
        'gassan_toda': '月山富田城',
        'adachi_museum': '足立美術館',
        'matsue_horikawa': '松江城(堀川遊覽船)',
        'matsue': '松江城',
        'tamatsukuri_yu_jinja': '玉作湯神社',
        'tamatsukuri_onsen': '玉造溫泉',
        'izumo_taisha': '出雲大社',
        'iwami_ginzan': '石見銀山',
        'hagi_lodging': '萩',
        'motonosumi_jinja': '元乃隅神社',
        'hagi': '萩城',
        'shin_yamaguchi_lodging': '新山口',
        'akiyoshidai_akiyoshido': '秋吉台・秋芳洞',
        'rurikoji_pagoda': '瑠璃光寺五重塔',
        'hofu_tenmangu': '防府天滿宮',
        'kintaikyo': '錦帶橋',
        'iwakuni': '岩國城',
        'itsukushima_jinja': '嚴島神社',
        'daishoin': '宮島彌山大聖院',
        'hiroshima_lodging_d8': '廣島',
        'hiroshima': '廣島城',
        'hiroshima_gokoku_jinja': '廣島護國神社',
        'pokemon_center_hiroshima': '寶可夢中心',
        'atomic_bomb_dome': '原爆圓頂',
        'hiroshima_peace_museum': '和平紀念資料館',
        'hiroshima_jinseki': '鉄板焼 神石',
        'hiroshima_lodging_d9': '廣島',
        'yoshida_koriyama': '吉田郡山城',
        'fukuyama': '福山城',
        'tomonoura_onsen': '鞆之浦溫泉',
        'bitchu_matsuyama': '備中松山城',
        'kibitsuhiko_jinja': '吉備津彥神社',
        'kibitsu_jinja': '吉備津神社',
        'okayama_hotel_d11': '岡山'
    }
    
    for sid in spot_ids:
        if sid in custom_names:
            c = custom_names[sid]
        elif sid in spots_map:
            c = spots_map[sid].get('name', '')
            c = re.sub(r'\(.*?\)', '', c).strip()
        else:
            continue
        if c and (not names or names[-1] != c):
            if c in ['新山口', '廣島', '岡山', '米子', '鳥取'] and names and names[-1] == c:
                continue
            names.append(c)
    
    if d_num == 14 and not names:
        names = ['岡山桃太郎機場(返台)']
    elif d_num == 14:
        names[-1] = '岡山桃太郎機場(返台)'
        
    return f"D{d_num}:" + ".".join(names)

LOCAL_GOURMET = {
    1: {
        "region": "岡山站前・表町商圈",
        "items": [
            ("🥩 岡山千屋牛", "日本最古老蔓牛血統，推薦格蘭比亞飯店2F「鉄板焼 備彩」品嚐頂級黑毛和牛鐵板燒。"),
            ("🍑 岡山果實甜點", "岡山白桃霜淇淋、晴王麝香葡萄夾心餅乾與新鮮水果塔。"),
            ("🍤 岡山庶民名物", "岡山名物蝦仁黑炒飯（えびめし）、散壽司（ばらずし）。")
        ],
        "tips": "抵達首夜推薦直奔岡山千屋牛鐵板燒享受頂級和牛晚宴，餐後漫步 AEON Mall 採買。"
    },
    2: {
        "region": "吉備總社・美作津山・鳥取",
        "items": [
            ("🍜 津山名物牛內臟炒烏龍麵", "津山ホルモンうどん・鶴山公園下町B級美食冠軍，獨門特製味噌醬汁炒新鮮牛內臟，香濃多汁。"),
            ("🥩 備中千屋牛特上牛排", "吉備高原產地直送頂級和牛。"),
            ("🍲 鳥取牛骨拉麵", "鳥取牛骨ラーメン・牛骨慢火清燉清香甘甜湯頭，鳥取在地靈魂名產。")
        ],
        "tips": "中午於津山城下町品嚐熱騰騰鐵板牛內臟烏龍麵補足體力；晚間抵達鳥取站前品嚐牛骨拉麵。"
    },
    3: {
        "region": "鳥取市區・但馬城崎溫泉",
        "items": [
            ("🍲 鳥取和牛元祖涮涮鍋", "鳥取民藝名店「たくみ割烹店」・日本牛肉涮涮鍋發祥店，極品和牛油脂入口即化。"),
            ("🍦 砂丘會館二十世紀梨冰淇淋", "純天然二十世紀梨原汁製作，酸甜清爽解膩。"),
            ("🦀 冬季城崎松葉蟹會席", "津居山港直送活蟹・刺身、烤蟹、甲羅燒與松葉蟹火鍋全席，鮮甜飽滿。"),
            ("🥩 但馬牛牛排與可樂餅", "城崎溫泉街但馬牛專門店熱騰騰現炸名物。")
        ],
        "tips": "中午在鳥取必吃たくみ割烹店和牛涮涮鍋；晚間入住城崎溫泉享用最頂級的松葉蟹全席！"
    },
    4: {
        "region": "丹後天橋立・伊根舟屋・米子",
        "items": [
            ("🐟 伊根冬季油寒鰤魚", "伊根名物「鰤しゃぶ」寒鰤魚涮涮鍋・冬季日本海油脂豐潤、肉質鮮甜緊實。"),
            ("🍢 宮津黑竹輪與海鮮丼", "現烤黑竹輪（黒ちくわ）與丹後沿岸當令海鮮丼。"),
            ("🦀 米子境港直送海鮮割烹", "米子站前「日本料理 荒磯」或「太平記」品嚐紅楚蟹、活白烏賊、鬼蜆湯與海鮮散壽司。")
        ],
        "tips": "中午在伊根舟屋海邊餐廳享用海鮮午餐；晚間於米子站前品嚐日本海直送海鮮割烹！"
    },
    5: {
        "region": "石見津和野・米子皆生溫泉",
        "items": [
            ("🫔 津和野名物源氏卷", "薄烤蛋糕皮包裹細緻紅豆沙・表參道山田竹風軒現烤熱騰騰名菓。"),
            ("🐟 高津川清流香魚料理", "日本第一清流高津川天然香魚鹽烤與山菜天婦羅定食。"),
            ("🦀 皆生溫泉日本海會席", "境港松葉蟹、日本海喉黑魚（ノドグロ）鹽烤、大山地雞與海濱御前料理。")
        ],
        "tips": "津和野殿町通散策時來份現做熱源氏卷配熱茶；晚間於海濱旅館享受含鹽泉質與奢華海味。"
    },
    6: {
        "region": "安來・出雲松江・玉造溫泉",
        "items": [
            ("🐟 宍道湖七珍名物料理", "宍道湖名物大和蜆味噌湯（しじみ汁）、鱸魚白杉燒、鰻魚。"),
            ("🍜 出雲割子蕎麥麵", "三段圓漆盒盛裝・八雲庵或神代蕎麥，淋上濃口柴魚醬汁與蔥花海苔。"),
            ("🍵 松江三大銘菓與抹茶", "山川、若草・感受松平不昧公城下町傳承數百年的茶道點心文化。"),
            ("🥩 島根和牛美肌會席", "玉造溫泉旅館頂級晚宴，搭配島根美酒。")
        ],
        "tips": "登松江城後至堀川周邊享用正宗割子蕎麥麵；晚間於玉造溫泉旅館享用宍道湖名物與島根和牛。"
    },
    7: {
        "region": "出雲大社・石見大森・長州萩市",
        "items": [
            ("🍜 出雲大社門前蕎麥麵", "神門通「八雲」或「荒木屋」割子蕎麥麵與出雲ぜんざい紅豆年糕湯。"),
            ("☕ 石見銀山群言堂旬菜料理", "大森街區古民家健康旬彩定食、銀山古法烘焙咖啡。"),
            ("🥩 萩見蘭牛極上和牛", "天然紀念物見島牛與荷斯登牛交配之稀少夢幻和牛・網燒牛排「玄」或「見蘭」。"),
            ("🦑 萩仙崎特鮮活烏賊", "長門仙崎港直送活透抽刺身，透明清甜彈牙。")
        ],
        "tips": "參拜出雲大社後於神門通吃割子蕎麥；抵達萩市後務必尋訪見蘭牛與日本海新鮮海產。"
    },
    8: {
        "region": "長門仙崎・北長門・新山口",
        "items": [
            ("🥩 萩網燒見蘭牛", "道の駅 萩往還「玄」牛排丼、見蘭肉餅、長州黑柏地雞。"),
            ("🍊 萩名產夏蜜柑甜點", "夏蜜柑霜淇淋、夏蜜柑果凍、夏蜜柑丸漬。"),
            ("🐡 山口南風河豚料理", "新山口站前名店品嚐河豚刺身（ふぐ刺し）、炸河豚、河豚鰭酒。"),
            ("🍢 長門名物烤雞肉串", "仙崎深川養雞場直送炭火烤雞串（やきとり）。")
        ],
        "tips": "下午於萩城下町或道の駅享用見蘭牛與夏蜜柑甜點；晚間於新山口站前居酒屋品嚐山口河豚與獺祭地酒。"
    },
    9: {
        "region": "秋吉台・山口市・防府天滿宮",
        "items": [
            ("🍜 秋吉台牛蒡烏龍麵", "秋芳洞名物ごぼううどん・香酥牛蒡絲天婦羅搭配清爽柴魚高湯。"),
            ("🍡 山口名物外郎糕", "山口ういろう・蕨粉揉製，口感比名古屋外郎更Ｑ彈滑潤，豆沙與抹茶絕品。"),
            ("🍵 防府天滿宮梅枝餅", "表參道名物現烤梅花烤年糕餅，外脆內軟。"),
            ("🥢 山口瓦片燒蕎麥麵", "元祖瓦そば・熱燙瓦片上盛放茶蕎麥、牛肉絲、錦絲蛋，沾特製熱醬汁。")
        ],
        "tips": "秋芳洞出洞後可在觀光商店街吃碗熱牛蒡烏龍麵；防府天滿宮參拜後品嚐梅枝餅；晚間於新山口吃瓦片蕎麥麵。"
    },
    10: {
        "region": "周防岩國・安藝宮島・廣島",
        "items": [
            ("🍱 錦帶橋岩國壽司", "殿樣壽司・五彩木押壽司，蓮藕、蛋絲、香菇層層堆疊色彩華麗。"),
            ("🦪 宮島烤牡蠣與炸牡蠣", "宮島名店「牡蠣屋」或「沖野水產」・瀨戶內海肥美大牡蠣，多汁甘甜。"),
            ("🍚 宮島百年穴子飯", "宮島口「うえの 穴子めし」・秘傳醬汁蒲燒星鰻飯，米飯吸滿高湯香氣。"),
            ("🍁 宮島紅葉饅頭", "もみじ饅頭・生紅葉饅頭、現炸紅葉饅頭（紅葉堂）。")
        ],
        "tips": "中午在岩國錦帶橋旁品嚐岩國壽司；下午宮島表參道吃現烤牡蠣、炸紅葉饅頭與穴子飯！"
    },
    11: {
        "region": "廣島市區・八丁堀・紙屋町",
        "items": [
            ("🥩 頂級廣島牛鐵板燒晚宴", "格蘭比亞 21F「鉄板焼 神石」・品嚐最高等級廣島和牛與瀨戶內鮑魚。"),
            ("🥞 廣島本通正宗廣島燒", "お好み焼き・長田屋、みっちゃん 或 麗ちゃん，層層疊上高麗菜、豬肉、中華炒麵。"),
            ("🍜 廣島沾麵", "廣島つけ麺・ばくだん屋，冰鎮辛香辣沾醬配冷麵，酸辣開胃過癮。"),
            ("🥐 八天堂廣島限定奶油麵包", "入口即化的濃郁卡士達與鮮奶油。")
        ],
        "tips": "中午在紙屋町/本通吃現煎廣島燒；晚間在飯店高空餐廳「鉄板焼 神石」享受極致和牛鐵板燒盛宴。"
    },
    12: {
        "region": "安藝高田・備後福山・鞆之浦",
        "items": [
            ("🍜 安藝高田夜叉烏龍麵", "夜叉うどん・在地名物，秘傳辣味噌、特產蔥與黑豚肉片，暖胃過癮。"),
            ("🐟 鞆之浦名物鯛魚料理", "鯛魚茶泡飯（鯛茶漬け）、鯛魚飯（鯛めし）、鯛魚荒焚き，肉質甘甜緊實。"),
            ("🍶 鞆之浦十六味保命酒", "江戶傳承至今之漢方本草藥酒，可試飲保命酒飴與保命酒冰淇淋。"),
            ("🐟 瀨戶內海小魚料理", "ガスエビ、小イワシ天婦羅酥脆美味。")
        ],
        "tips": "中午探訪郡山城時可在安藝高田品嚐夜叉烏龍麵；傍晚抵達鞆之浦享用海景溫泉飯店的經典鯛魚全席晚宴！"
    },
    13: {
        "region": "備中高梁・吉備路・岡山",
        "items": [
            ("🍝 高梁名物印地安番茄炒麵", "インディアントマト焼そば・咖哩風味炒麵淋上高梁完熟番茄肉醬。"),
            ("🍱 岡山味司野村多蜜豬排丼", "デミカツ丼・炸豬排淋上法式濃郁多蜜醬汁，岡山獨特代表名物。"),
            ("🦪 備前日生牡蠣文字燒", "カキオコ・鐵板上鋪滿肥美鮮蚵，鮮香撲鼻。"),
            ("🍑 岡山白桃果凍與吉備糰子", "名產街伴手禮首選。")
        ],
        "tips": "高梁城下町可品嚐印地安番茄炒麵；返回岡山後於市區品嚐名物醬汁豬排丼，並進行最終採買！"
    },
    14: {
        "region": "岡山市區・岡山桃太郎機場",
        "items": [
            ("🍡 岡山大手饅頭", "大手まんぢゅう・天保年間創立，薄皮紅豆甘酒名菓。"),
            ("🍡 廣榮堂吉備糰子", "元祖きびだんご・桃太郎故事傳說伴手禮。"),
            ("🍪 岡山白桃夾心法蘭酥與晴王麝香葡萄巧克力", "機場免稅店人氣第一名產。"),
            ("🍱 機場限定和牛便當與岡山特產美食", "登機前最後犒賞。")
        ],
        "tips": "離境前於岡山站一番街與桃太郎機場出境免稅店採買齊全岡山與山陰山陽特產銘菓，滿載而歸！"
    }
}

def build_full_day_schedule(d_num, itinerary, spots_map, fares_list):

    day = next(d for d in itinerary if d.get('dayNumber') == d_num)
    spot_ids = day.get('spotIds', [])
    day_spots = [spots_map[sid] for sid in spot_ids if sid in spots_map]
    day_fares = [f for f in fares_list if f.get('dayNumber') == d_num]
    att_dict = day.get('transitAttachments', {})
    gm = LOCAL_GOURMET.get(d_num, {})
    gm_region = gm.get('region', '當地名產')
    gm_title = f"🍽️ {gm_region} 特色美食與名物料理"
    if isinstance(gm.get('items'), list):
        gm_items_str = " ｜ ".join([f"{t}（{d}）" for t, d in gm['items']])
    else:
        gm_items_str = str(gm.get('items', ''))
    gm_desc = f"必吃名物：{gm_items_str}\n💡 賞味叮嚀：{gm.get('tips', '')}"
    d_notes = day.get('notes', '')
    
    # We want a unified list of steps for the day:
    # Each step: (phase_name, item_title, category_mode, cost_pass, details_tips)
    steps = []
    
    # Helper to format castle stamp
    def get_spot_stamp_info(sp):
        is_castle = (sp.get('type') == 'castle')
        stamp = sp.get('stampLocation', '')
        open_h = sp.get('openHours', '全年無休')
        closed = sp.get('closedDayText', '')
        transit = sp.get('transitInfo', '')
        
        parts = []
        if is_castle and stamp and stamp != '無':
            parts.append(f"🏯 百名城印章：{stamp}")
        parts.append(f"⏰ 開放時間：{open_h}" + (f" (休：{closed})" if closed else ""))
        if transit:
            parts.append(f"🚶 接駁提示：{transit[:70]}")
        return "\n".join(parts)
        
    def get_fare_cost(fare):
        amt = fare.get('amount', 0)
        p = fare.get('passCovered', False)
        return f"¥{amt:,}\n(🎫 Pass免費)" if p else f"¥{amt:,}\n(⚠️ 需自費)"

    # Build sequence based on Day
    if d_num == 1:
        steps.append(("14:30~15:05", "抵達岡山桃太郎機場 (OKJ)", "班機抵達", "-", "搭乘台灣虎航 IT214 降落，出關提領行李"))
        if day_fares:
            steps.append(("15:30~16:05", day_fares[0]['title'], day_fares[0]['transitTypeName'], get_fare_cost(day_fares[0]), day_fares[0]['notes']))
        steps.append(("16:30~17:30", "VIA INN 岡山 (岡山站直結)", "飯店入住", "-", "出站直通飯店辦理 Check-in，卸下行李"))
        steps.append(("18:00~19:30", gm_title, "當地美食", "自費", gm_desc))
        steps.append(("19:30~21:00", "AEON Mall Okayama (西日本最大級)", "逛街採買", "-", "漫步350間店鋪商圈採買藥妝物資，儲備體力"))

    elif d_num == 2:
        if len(day_fares) >= 2:
            steps.append(("07:30~08:10", day_fares[0]['title'], day_fares[0]['transitTypeName'], get_fare_cost(day_fares[0]), day_fares[0]['notes']))
            steps.append(("08:15~08:45", day_fares[1]['title'], day_fares[1]['transitTypeName'], get_fare_cost(day_fares[1]), day_fares[1]['notes']))
        sp_kinojo = next((s for s in day_spots if s.get('id') == 'kinojo'), None)
        if sp_kinojo:
            steps.append(("08:45~10:45", sp_kinojo['name'], "百名城參訪", "-", get_spot_stamp_info(sp_kinojo)))
        steps.append(("10:50~11:45", "鬼ノ城 ➔ 總社駅 ➔ 岡山駅 (預約計程車下山 ➔ JR 回岡山)", "電車", "¥420\n(⚠️ 需自費)", "非 Pass 天數・山頂搭乘約定計程車下山返回總社站 (約 25 分)，轉乘 JR 伯備線或吉備線普通車約 30 分返回岡山站東口，準備享用午餐"))
        steps.append(("12:00~13:00", gm_title, "當地美食", "自費", gm_desc))
        steps.append(("13:05~14:15", "岡山駅 ➔ 津山駅 (JR 津山線 快速ことぶき / 普通車 往津山城)", "電車", "¥1,170\n(⚠️ 需自費)", "非 Pass 天數・於岡山站搭乘 JR 津山線快速列車 (快速ことぶき) 直達津山站，車程約 1 小時 10 分，刷 IC 卡或購票乘車"))
        if len(day_fares) >= 4:
            steps.append(("14:15~14:30", day_fares[3]['title'], day_fares[3]['transitTypeName'], get_fare_cost(day_fares[3]), day_fares[3]['notes']))
        sp_tsuyama = next((s for s in day_spots if s.get('id') == 'tsuyama'), None)
        if sp_tsuyama:
            steps.append(("14:30~16:30", sp_tsuyama['name'], "百名城參訪", "-", get_spot_stamp_info(sp_tsuyama)))
        if len(day_fares) >= 5:
            steps.append(("16:38~18:00", day_fares[4]['title'], day_fares[4]['transitTypeName'], get_fare_cost(day_fares[4]), day_fares[4]['notes']))
        steps.append(("18:30~", "APA飯店 鳥取車站前 ＆ 市區晚餐", "住宿 Check-in", "-", "出站徒步進飯店，晚間品嚐鳥取牛骨拉麵"))

    elif d_num == 3:
        if len(day_fares) >= 1:
            steps.append(("07:35~07:42", day_fares[0]['title'], day_fares[0]['transitTypeName'], get_fare_cost(day_fares[0]), day_fares[0]['notes']))
        sp_tottori = next((s for s in day_spots if s.get('id') == 'tottori'), None)
        if sp_tottori:
            steps.append(("07:45~09:40", sp_tottori['name'], "百名城參訪", "-", get_spot_stamp_info(sp_tottori)))
        if len(day_fares) >= 3:
            steps.append(("09:47~10:02", day_fares[1]['title'], day_fares[1]['transitTypeName'], get_fare_cost(day_fares[1]), day_fares[1]['notes']))
            steps.append(("10:05~10:15", day_fares[2]['title'], day_fares[2]['transitTypeName'], get_fare_cost(day_fares[2]), day_fares[2]['notes']))
        sp_sakyu = next((s for s in day_spots if s.get('id') == 'tottori_sakyu'), None)
        if sp_sakyu:
            steps.append(("10:15~11:15", sp_sakyu['name'], "景點參訪", "-", get_spot_stamp_info(sp_sakyu)))
        if len(day_fares) >= 4:
            steps.append(("11:20~11:42", day_fares[3]['title'], day_fares[3]['transitTypeName'], get_fare_cost(day_fares[3]), day_fares[3]['notes']))
        steps.append(("11:45~13:10", gm_title, "當地美食", "自費", gm_desc))
        if len(day_fares) >= 6:
            steps.append(("13:33~15:28", day_fares[5]['title'], day_fares[5]['transitTypeName'], get_fare_cost(day_fares[5]), day_fares[5]['notes']))
        steps.append(("15:30~", "城崎溫泉 和風旅館（住宿 Check-in ＆ 松葉蟹晚宴）", "夜宿溫泉", "-", "換穿浴衣漫步七大外湯街，享用冬季津居山港直送活松葉蟹全席"))

    elif d_num == 4:
        steps.append(("07:06~08:45", "城崎溫泉駅 ➔ 天橋立駅 (JR ➔ 豐岡轉乘 ➔ 京都丹後鐵道)", "電車/快速", "Pass包含", "07:06 城崎溫泉發車 ➔ 豐岡站轉乘 ➔ 08:45 抵達天橋立駅 (持 JR Pass 或丹後天橋立伊根周遊券包含此段)"))
        steps.append(("09:00~09:12", "天橋立棧橋 ➔ 一之宮碼頭 (天橋立觀光船・橫渡松林沙洲)", "觀光遊覽船", "Pass包含", "天橋立站步行5分至天橋立棧橋搭09:00開航觀光船 ➔ 09:12 抵一之宮碼頭 (丹後天橋立伊根周遊券免費搭乘)"))
        steps.append(("09:15~11:15", "天橋立傘松公園 (登山纜車・單人吊椅 / 昇龍觀 股のぞき)", "登山纜車/吊椅", "Pass包含", "一宮碼頭步行5分至府中站，搭09:30纜車登頂傘松公園跨下倒看「昇龍觀」一線天絕景；需提早於11:15前搭纜車下山"))
        steps.append(("11:15~11:24", "えとせとら松屋 (黑竹輪) ＆ 元伊勢籠神社參拜", "參道漫步", "自費", "下山品嚐松屋現烤名物黑竹輪，順訪元伊勢籠神社，隨後至神社前公車站牌候車"))
        steps.append(("11:24~11:50", "神社前 (籠神社前) ➔ 伊根灣巡游觀光船日出站 (丹海巴士)", "路線公車", "Pass包含", "11:24 於神社前搭乘丹海巴士 ➔ 11:50 抵達「伊根湾めぐり・日出」站 (憑周遊券免費搭乘)"))
        steps.append(("12:00~12:25", "伊根灣舟屋遊覽船 (海上巡航・餵海鷗・欣賞230棟水上舟屋)", "觀光遊覽船", "Pass包含", "搭乘12:00伊根灣遊覽船環灣巡航25分鐘，甲板餵海鷗、飽覽臨海舟屋群 (憑周遊券免費搭乘)"))
        steps.append(("12:25~13:05", "伊根灣周邊散策與休憩", "聚落漫步", "自費", "於伊根日出棧橋周邊賞景散策拍照，13:00 前往日出站公車站牌準備候車"))
        steps.append(("13:09~14:02", "伊根灣日出站 ➔ 天橋立駅 (丹海路線巴士 返站)", "路線公車", "Pass包含", "13:09 於日出站搭乘丹海巴士回程 ➔ 14:02 準時抵達天橋立車站"))
        steps.append(("14:05~15:15", "はしだて茶屋 (Hashidate Chaya) 享用名物花蛤丼 ＆ 迴旋橋散策", "名物美食＆散策", "自費", "出站步行9分穿過迴旋橋，享用天橋立名物「あさり丼(花蛤丼)」午餐與日式甜品，餐後步行回車站"))
        steps.append(("15:21~21:54", "天橋立駅 ➔ 米子駅 (丹鐵宮豐線 ➔ 豐岡 ➔ JR山陰本線 ➔ 濱坂 ➔ 鳥取 ➔ 特急スーパーまつかぜ13號)", "特急/快速轉乘", "Pass免費", "15:21 天橋立發 ➔ 16:36 抵豐岡 17:29 發 ➔ 18:36 抵濱坂 19:07 發 ➔ 20:02 抵鳥取 20:50 發 ➔ 21:54 抵達米子駅！(JR Pass 全額免費，現省 5,480 円)"))
        steps.append(("22:00~", "米子站前 Hotel Harvest In Yonago 入住 ＆ 站前宵夜居酒屋 (居酒屋 太平記)", "夜宿飯店", "-", "出站徒步 1 分鐘進飯店 Check-in，晚間可步行至米子站前營業至 23:00 的「居酒屋 太平記」小酌享用境港當季海鮮、大山地雞串燒"))

    elif d_num == 5:
        if len(day_fares) >= 1:
            steps.append(("06:00~09:15", day_fares[0]['title'], day_fares[0]['transitTypeName'], get_fare_cost(day_fares[0]), day_fares[0]['notes']))
        if len(day_fares) >= 2:
            steps.append(("09:20~09:30", day_fares[1]['title'], day_fares[1]['transitTypeName'], get_fare_cost(day_fares[1]), day_fares[1]['notes']))
        sp_tsuwano = next((s for s in day_spots if s.get('id') == 'tsuwano'), None)
        if sp_tsuwano:
            steps.append(("09:30~11:30", sp_tsuwano['name'], "百名城參訪", "-", get_spot_stamp_info(sp_tsuwano)))
        steps.append(("11:30~12:30", gm_title, "當地美食", "自費", gm_desc))
        sp_taikodani = next((s for s in day_spots if s.get('id') == 'taikodani_inari'), None)
        if sp_taikodani:
            steps.append(("12:30~13:30", sp_taikodani['name'], "神社參拜", "-", get_spot_stamp_info(sp_taikodani)))
        if len(day_fares) >= 4:
            steps.append(("14:15~17:00", day_fares[2]['title'], day_fares[2]['transitTypeName'], get_fare_cost(day_fares[2]), day_fares[2]['notes']))
            steps.append(("17:18~17:37", day_fares[3]['title'], day_fares[3]['transitTypeName'], get_fare_cost(day_fares[3]), day_fares[3]['notes']))
        steps.append(("18:00~", "皆生溫泉 海濱和風溫泉旅館（住宿 Check-in ＆ 日本海晚宴）", "夜宿溫泉", "-", "享受鹽化妝水神之湯，品嚐日本海境港松葉蟹與喉黑魚會席料理"))

    elif d_num == 6:
        if len(day_fares) >= 1:
            steps.append(("07:30~07:50", day_fares[0]['title'], day_fares[0]['transitTypeName'], get_fare_cost(day_fares[0]), day_fares[0]['notes']))
        if len(day_fares) >= 2:
            steps.append(("08:15~08:45", day_fares[1]['title'], day_fares[1]['transitTypeName'], get_fare_cost(day_fares[1]), day_fares[1]['notes']))
        sp_gassan = next((s for s in day_spots if s.get('id') == 'gassan_toda'), None)
        if sp_gassan:
            steps.append(("08:45~10:45", sp_gassan['name'], "百名城參訪", "-", get_spot_stamp_info(sp_gassan)))
        if len(day_fares) >= 4:
            steps.append(("10:55~11:15", day_fares[2]['title'], day_fares[2]['transitTypeName'], get_fare_cost(day_fares[2]), day_fares[2]['notes']))
        sp_adachi = next((s for s in day_spots if s.get('id') == 'adachi_museum'), None)
        if sp_adachi:
            steps.append(("11:15~12:45", sp_adachi['name'], "景點參訪", "-", get_spot_stamp_info(sp_adachi)))
        steps.append(("12:45~13:30", gm_title, "當地美食", "自費", gm_desc))
        if len(day_fares) >= 6:
            steps.append(("13:30~13:50", day_fares[3]['title'], day_fares[3]['transitTypeName'], get_fare_cost(day_fares[3]), day_fares[3]['notes']))
            steps.append(("14:00~14:25", day_fares[4]['title'], day_fares[4]['transitTypeName'], get_fare_cost(day_fares[4]), day_fares[4]['notes']))
            steps.append(("14:30~14:45", day_fares[5]['title'], day_fares[5]['transitTypeName'], get_fare_cost(day_fares[5]), day_fares[5]['notes']))
        sp_matsue = next((s for s in day_spots if s.get('id') == 'matsue'), None)
        if sp_matsue:
            steps.append(("14:45~16:45", sp_matsue['name'] + " (含堀川遊覽船)", "國寶名城參訪", "-", get_spot_stamp_info(sp_matsue)))
        if len(day_fares) >= 7:
            steps.append(("17:00~17:35", day_fares[6]['title'], day_fares[6]['transitTypeName'], get_fare_cost(day_fares[6]), day_fares[6]['notes']))
        steps.append(("18:00~", "玉造溫泉 和風旅館（住宿 Check-in ＆ 美肌晚宴）", "夜宿溫泉", "-", "入住日本第一美肌之湯，參拜玉作湯神社，享用島根和牛會席"))

    elif d_num == 7:
        if len(day_fares) >= 2:
            steps.append(("07:30~08:00", day_fares[0]['title'], day_fares[0]['transitTypeName'], get_fare_cost(day_fares[0]), day_fares[0]['notes']))
            steps.append(("08:05~08:35", day_fares[1]['title'], day_fares[1]['transitTypeName'], get_fare_cost(day_fares[1]), day_fares[1]['notes']))
        sp_izumo = next((s for s in day_spots if s.get('id') == 'izumo_taisha'), None)
        if sp_izumo:
            steps.append(("08:35~10:45", sp_izumo['name'], "神話大社參拜", "-", get_spot_stamp_info(sp_izumo)))
        steps.append(("11:00~12:00", gm_title, "當地美食", "自費", gm_desc))
        if len(day_fares) >= 5:
            steps.append(("12:15~13:00", day_fares[2]['title'], day_fares[2]['transitTypeName'], get_fare_cost(day_fares[2]), day_fares[2]['notes']))
            steps.append(("13:10~13:40", day_fares[3]['title'], day_fares[3]['transitTypeName'], get_fare_cost(day_fares[3]), day_fares[3]['notes']))
            steps.append(("13:45~14:00", day_fares[4]['title'], day_fares[4]['transitTypeName'], get_fare_cost(day_fares[4]), day_fares[4]['notes']))
        sp_iwami = next((s for s in day_spots if s.get('id') == 'iwami_ginzan'), None)
        if sp_iwami:
            steps.append(("14:00~16:30", sp_iwami['name'], "世界遺產巡禮", "-", get_spot_stamp_info(sp_iwami)))
        if len(day_fares) >= 7:
            steps.append(("17:15~20:27", day_fares[5]['title'], day_fares[5]['transitTypeName'], get_fare_cost(day_fares[5]), day_fares[5]['notes']))
            steps.append(("20:30~20:45", day_fares[6]['title'], day_fares[6]['transitTypeName'], get_fare_cost(day_fares[6]), day_fares[6]['notes']))
        steps.append(("21:00~", "萩市區飯店 / 萩溫泉旅館（住宿 Check-in）", "夜宿長州", "-", "入住維新歷史之城萩市，品嚐見蘭牛宵夜與溫泉休養"))

    elif d_num == 8:
        if len(day_fares) >= 2:
            steps.append(("07:45~08:20", day_fares[0]['title'], day_fares[0]['transitTypeName'], get_fare_cost(day_fares[0]), day_fares[0]['notes']))
            steps.append(("08:25~08:45", day_fares[1]['title'], day_fares[1]['transitTypeName'], get_fare_cost(day_fares[1]), day_fares[1]['notes']))
        sp_motonosumi = next((s for s in day_spots if s.get('id') == 'motonosumi_jinja'), None)
        if sp_motonosumi:
            steps.append(("08:45~10:15", sp_motonosumi['name'], "絕景神社參拜", "-", get_spot_stamp_info(sp_motonosumi)))
        if len(day_fares) >= 4:
            steps.append(("10:30~11:15", day_fares[2]['title'], day_fares[2]['transitTypeName'], get_fare_cost(day_fares[2]), day_fares[2]['notes']))
            steps.append(("11:30~11:45", day_fares[3]['title'], day_fares[3]['transitTypeName'], get_fare_cost(day_fares[3]), day_fares[3]['notes']))
        sp_hagi = next((s for s in day_spots if s.get('id') == 'hagi'), None)
        if sp_hagi:
            steps.append(("11:45~13:45", sp_hagi['name'], "百名城參訪", "-", get_spot_stamp_info(sp_hagi)))
        steps.append(("14:00~15:00", gm_title, "當地美食", "自費", gm_desc))
        if len(day_fares) >= 5:
            steps.append(("16:00~17:15", day_fares[4]['title'], day_fares[4]['transitTypeName'], get_fare_cost(day_fares[4]), day_fares[4]['notes']))
        steps.append(("17:30~", "Comfort Hotel Shin Yamaguchi ＆ 山口河豚晚宴", "住宿 Check-in", "-", "入住新山口站前連住第1晚，晚間居酒屋享用山口河豚刺身與獺祭"))

    elif d_num == 9:
        if len(day_fares) >= 1:
            steps.append(("07:35~08:15", day_fares[0]['title'], day_fares[0]['transitTypeName'], get_fare_cost(day_fares[0]), day_fares[0]['notes']))
        sp_akiyoshi = next((s for s in day_spots if s.get('id') == 'akiyoshidai_akiyoshido'), None)
        if sp_akiyoshi:
            steps.append(("08:15~10:30", sp_akiyoshi['name'], "自然奇景探索", "-", get_spot_stamp_info(sp_akiyoshi)))
        if len(day_fares) >= 2:
            steps.append(("10:40~11:38", day_fares[1]['title'], day_fares[1]['transitTypeName'], get_fare_cost(day_fares[1]), day_fares[1]['notes']))
        sp_rurikoji = next((s for s in day_spots if s.get('id') == 'rurikoji_pagoda'), None)
        if sp_rurikoji:
            steps.append(("11:45~13:15", sp_rurikoji['name'], "國寶佛塔參訪", "-", get_spot_stamp_info(sp_rurikoji)))
        steps.append(("13:15~13:30", gm_title, "當地美食", "自費", gm_desc))
        if len(day_fares) >= 4:
            steps.append(("13:28~14:12", day_fares[2]['title'], day_fares[2]['transitTypeName'], get_fare_cost(day_fares[2]), day_fares[2]['notes']))
        sp_hofu = next((s for s in day_spots if s.get('id') == 'hofu_tenmangu'), None)
        if sp_hofu:
            steps.append(("14:30~16:00", sp_hofu['name'], "三大天神參拜", "-", get_spot_stamp_info(sp_hofu)))
        if len(day_fares) >= 4:
            steps.append(("16:30~16:45", day_fares[3]['title'], day_fares[3]['transitTypeName'], get_fare_cost(day_fares[3]), day_fares[3]['notes']))
        steps.append(("17:00~", "Comfort Hotel Shin Yamaguchi ＆ 瓦片燒蕎麥晚宴", "住宿連住第2晚", "-", "返回新山口站前，晚間品嚐熱燙瓦片茶蕎麥麵與地產美食"))

    elif d_num == 10:
        if len(day_fares) >= 3:
            steps.append(("07:45~08:13", day_fares[0]['title'], day_fares[0]['transitTypeName'], get_fare_cost(day_fares[0]), day_fares[0]['notes']))
            steps.append(("08:20~08:35", day_fares[1]['title'], day_fares[1]['transitTypeName'], get_fare_cost(day_fares[1]), day_fares[1]['notes']))
        sp_kintai = next((s for s in day_spots if s.get('id') == 'kintaikyo'), None)
        if sp_kintai:
            steps.append(("08:35~09:15", sp_kintai['name'], "名橋漫步", "-", get_spot_stamp_info(sp_kintai)))
        if len(day_fares) >= 3:
            steps.append(("09:15~09:20", day_fares[2]['title'], day_fares[2]['transitTypeName'], get_fare_cost(day_fares[2]), day_fares[2]['notes']))
        sp_iwakuni = next((s for s in day_spots if s.get('id') == 'iwakuni'), None)
        if sp_iwakuni:
            steps.append(("09:20~11:00", sp_iwakuni['name'], "百名城參訪", "-", get_spot_stamp_info(sp_iwakuni)))
        steps.append(("11:15~12:00", gm_title, "當地美食", "自費", gm_desc))
        if len(day_fares) >= 6:
            steps.append(("12:10~12:30", day_fares[3]['title'], day_fares[3]['transitTypeName'], get_fare_cost(day_fares[3]), day_fares[3]['notes']))
            steps.append(("12:40~13:03", day_fares[4]['title'], day_fares[4]['transitTypeName'], get_fare_cost(day_fares[4]), day_fares[4]['notes']))
            steps.append(("13:10~13:20", day_fares[5]['title'], day_fares[5]['transitTypeName'], get_fare_cost(day_fares[5]), day_fares[5]['notes']))
        sp_itsukushima = next((s for s in day_spots if s.get('id') == 'itsukushima_jinja'), None)
        sp_daishoin = next((s for s in day_spots if s.get('id') == 'daishoin'), None)
        if sp_itsukushima:
            steps.append(("13:20~15:30", sp_itsukushima['name'] + " (海上鳥居)", "世界遺產巡禮", "-", get_spot_stamp_info(sp_itsukushima)))
        if sp_daishoin:
            steps.append(("15:30~16:45", sp_daishoin['name'], "古剎巡禮", "-", get_spot_stamp_info(sp_daishoin)))
        if len(day_fares) >= 7:
            steps.append(("17:15~17:45", day_fares[6]['title'], day_fares[6]['transitTypeName'], get_fare_cost(day_fares[6]), day_fares[6]['notes']))
        steps.append(("18:00~", "Via Inn Hiroshima Shinkansenguchi ＆ 宮島牡蠣晚宴", "住宿 Check-in", "-", "入住廣島站直結飯店連住第1晚，晚間品嚐廣島燒或烤牡蠣"))

    elif d_num == 11:
        if day_fares:
            steps.append(("08:30~09:00", day_fares[0]['title'], day_fares[0]['transitTypeName'], get_fare_cost(day_fares[0]), day_fares[0]['notes']))
        sp_hiroshima = next((s for s in day_spots if s.get('id') == 'hiroshima'), None)
        if sp_hiroshima:
            steps.append(("09:00~10:30", sp_hiroshima['name'] + " ＆ 護國神社", "百名城參訪", "-", get_spot_stamp_info(sp_hiroshima)))
        sp_poke = next((s for s in day_spots if s.get('id') == 'pokemon_center_hiroshima'), None)
        if sp_poke:
            steps.append(("10:45~11:45", sp_poke['name'], "逛街採買", "-", "參觀 SOGO 6F 寶可夢中心選購限定周邊"))
        steps.append(("12:00~13:00", gm_title, "當地美食", "自費", gm_desc))
        sp_dome = next((s for s in day_spots if s.get('id') == 'atomic_bomb_dome'), None)
        sp_museum = next((s for s in day_spots if s.get('id') == 'hiroshima_peace_museum'), None)
        if sp_dome or sp_museum:
            steps.append(("13:15~16:00", "原爆圓頂 ＆ 和平紀念資料館", "和平文化巡禮", "-", "悼念世界和平歷史，走訪和平紀念公園"))
        steps.append(("18:00~20:00", "格蘭比亞 21F「鉄板焼 神石」廣島牛晚宴", "頂級美食晚宴", "自費", "享用頂級廣島和牛、鮑魚鐵板燒高空夜景晚宴，宿廣島連住第2晚"))

    elif d_num == 12:
        if len(day_fares) >= 1:
            steps.append(("08:00~09:15", day_fares[0]['title'], day_fares[0]['transitTypeName'], get_fare_cost(day_fares[0]), day_fares[0]['notes']))
        sp_koriyama = next((s for s in day_spots if s.get('id') == 'yoshida_koriyama'), None)
        if sp_koriyama:
            steps.append(("09:15~12:15", sp_koriyama['name'], "百名城參訪", "-", get_spot_stamp_info(sp_koriyama)))
        steps.append(("12:30~13:30", gm_title, "當地美食", "自費", gm_desc))
        if len(day_fares) >= 2:
            steps.append(("14:00~15:00", day_fares[1]['title'], day_fares[1]['transitTypeName'], get_fare_cost(day_fares[1]), day_fares[1]['notes']))
        sp_fukuyama = next((s for s in day_spots if s.get('id') == 'fukuyama'), None)
        if sp_fukuyama:
            steps.append(("15:00~16:30", sp_fukuyama['name'], "百名城參訪", "-", get_spot_stamp_info(sp_fukuyama)))
        if len(day_fares) >= 3:
            steps.append(("16:45~17:15", day_fares[2]['title'], day_fares[2]['transitTypeName'], get_fare_cost(day_fares[2]), day_fares[2]['notes']))
        steps.append(("17:30~", "鞆之浦溫泉 遠景之宿（海景溫泉 ＆ 鯛魚全席）", "夜宿溫泉", "-", "入住海景溫泉飯店，欣賞鞆之浦常夜燈港町夜景，品嚐鯛魚名物料理"))

    elif d_num == 13:
        if len(day_fares) >= 2:
            steps.append(("07:18~07:51", day_fares[0]['title'], day_fares[0]['transitTypeName'], get_fare_cost(day_fares[0]), day_fares[0]['notes']))
            steps.append(("08:00~08:16", day_fares[1]['title'], day_fares[1]['transitTypeName'], get_fare_cost(day_fares[1]), day_fares[1]['notes']))
        if len(day_fares) >= 4:
            steps.append(("08:29~09:49", day_fares[2]['title'], day_fares[2]['transitTypeName'], get_fare_cost(day_fares[2]), day_fares[2]['notes']))
            steps.append(("10:00~10:20", day_fares[3]['title'], day_fares[3]['transitTypeName'], get_fare_cost(day_fares[3]), day_fares[3]['notes']))
        sp_matsuyama = next((s for s in day_spots if s.get('id') == 'bitchu_matsuyama'), None)
        if sp_matsuyama:
            steps.append(("10:20~12:30", sp_matsuyama['name'], "現存山城參訪", "-", get_spot_stamp_info(sp_matsuyama)))
        steps.append(("12:45~13:30", gm_title, "當地美食", "自費", gm_desc))
        if len(day_fares) >= 6:
            steps.append(("13:45~14:30", day_fares[4]['title'], day_fares[4]['transitTypeName'], get_fare_cost(day_fares[4]), day_fares[4]['notes']))
        sp_kibitsuhiko = next((s for s in day_spots if s.get('id') == 'kibitsuhiko_jinja'), None)
        if sp_kibitsuhiko:
            steps.append(("14:30~15:20", sp_kibitsuhiko['name'], "桃太郎神話古社", "-", get_spot_stamp_info(sp_kibitsuhiko)))
        if len(day_fares) >= 6:
            steps.append(("15:25~15:30", day_fares[5]['title'], day_fares[5]['transitTypeName'], get_fare_cost(day_fares[5]), day_fares[5]['notes']))
        sp_kibitsu = next((s for s in day_spots if s.get('id') == 'kibitsu_jinja'), None)
        if sp_kibitsu:
            steps.append(("15:30~16:45", sp_kibitsu['name'], "國寶大迴廊神社", "-", get_spot_stamp_info(sp_kibitsu)))
        if len(day_fares) >= 7:
            steps.append(("16:55~17:12", day_fares[6]['title'], day_fares[6]['transitTypeName'], get_fare_cost(day_fares[6]), day_fares[6]['notes']))
        steps.append(("17:30~", "VIA INN 岡山 ＆ 多蜜醬汁豬排丼晚宴", "住宿 Check-in", "-", "入住岡山站直通飯店，前往味司野村吃名物豬排丼，最後採買伴手禮"))

    elif d_num == 14:
        steps.append(("09:00~11:30", "岡山市區漫遊 ＆ 車站一番街伴手禮採買", "伴手禮採買", "-", "購買大手饅頭、廣榮堂吉備糰子、岡山白桃果凍名產"))
        steps.append(("11:30~12:30", gm_title, "當地美食", "自費", gm_desc))
        if day_fares:
            steps.append(("12:40~13:10", day_fares[0]['title'], day_fares[0]['transitTypeName'], get_fare_cost(day_fares[0]), day_fares[0]['notes']))
        steps.append(("13:15~15:30", "岡山桃太郎機場 (OKJ) 報到返台", "航班登機", "-", "辦理台灣虎航 IT215 登機手續 (15:55 起飛 ➔ 17:40 抵桃園 T1)，滿載賦歸"))

    return steps


def create_booklet_word():
    # 讀取 JSON 資料
    with open('spots_data.json', 'r', encoding='utf-8') as f:
        data = json.load(f)

    spots_map = {s['id']: s for s in data.get('spots', [])}
    itinerary = data.get('itinerary', [])
    fares_list = data.get('fares', [])

    doc = Document()
    add_header_footer(doc)

    # ==========================================
    # 1. 封面 (Cover Page)
    # ==========================================
    p_pre = doc.add_paragraph()
    p_pre.paragraph_format.space_before = Pt(40)
    p_pre.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub0 = p_pre.add_run("🇯🇵 2027 嚴選壯遊隨身專屬手冊")
    format_run(r_sub0, size_pt=13, bold=True, color_rgb=(37, 99, 235))

    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(10)
    p_title.paragraph_format.space_after = Pt(12)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("山陰山陽 14 天縱走\n名城・名湯・絕景漫遊手冊")
    format_run(r_title, size_pt=24, bold=True, color_rgb=(30, 58, 138))

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(36)
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run("日本100名城巡禮 ｜ 三大名湯 ｜ 日本海極品松葉蟹 ｜ 世界遺產石見銀山")
    format_run(r_sub, size_pt=11, color_rgb=(100, 116, 139))

    # 封面資訊卡片 (表格)
    card_table = doc.add_table(rows=6, cols=2)
    card_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    card_table.autofit = False

    info_data = [
        ("📅 旅行日期", "2027 年 1 月 12 日 (二) ～ 1 月 25 日 (一) 共 14 天 13 夜"),
        ("👥 隨行成員", "獅子 & 方 (雙人自助旅行・全程大眾運輸)"),
        ("✈️ 國際航班", "台灣虎航直飛岡山\n去程 IT214：2027/01/12 台北(桃園) 11:30 ➔ 岡山 15:05\n回程 IT215：2027/01/25 岡山 15:55 ➔ 台北(桃園) 17:40"),
        ("🎫 交通周遊券", "JR 山陰&山陽鐵路周遊券 7日券\n【Day 3 ～ Day 9 正式啟用】含山陰特急/丹後鐵道/智頭急行指定席"),
        ("🏯 攻城重點", "鬼之城、津山城、鳥取城、月山富田城、松江城(國寶)、津和野城、萩城、岩國城、廣島城、吉田郡山城、福山城、備中松山城"),
        ("🆘 緊急聯絡", "日本報警：110 ｜ 救護火警：119\n外交部旅外急難求助：+886-800-085-095\n台北駐大阪經濟文化辦事處：+81-6-6227-8623")
    ]

    col_widths = [Inches(1.8), Inches(4.7)]
    for row_idx, (k, v) in enumerate(info_data):
        row = card_table.rows[row_idx]
        
        # 標題格
        c0 = row.cells[0]
        c0.width = col_widths[0]
        set_cell_background(c0, "F8FAFC")
        set_cell_margins(c0, top=140, bottom=140, left=180, right=140)
        set_cell_border(c0, bottom={'sz': 4, 'val': 'single', 'color': 'E2E8F0'},
                            top={'sz': 4, 'val': 'single', 'color': 'E2E8F0'},
                            left={'sz': 12, 'val': 'single', 'color': '2563EB'},
                            right={'sz': 4, 'val': 'single', 'color': 'E2E8F0'})
        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r0 = p0.add_run(k)
        format_run(r0, size_pt=9.5, bold=True, color_rgb=(30, 58, 138))

        # 內容格
        c1 = row.cells[1]
        c1.width = col_widths[1]
        set_cell_background(c1, "FFFFFF")
        set_cell_margins(c1, top=140, bottom=140, left=180, right=180)
        set_cell_border(c1, bottom={'sz': 4, 'val': 'single', 'color': 'E2E8F0'},
                            top={'sz': 4, 'val': 'single', 'color': 'E2E8F0'},
                            left={'sz': 4, 'val': 'single', 'color': 'E2E8F0'},
                            right={'sz': 4, 'val': 'single', 'color': 'E2E8F0'})
        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r1 = p1.add_run(v)
        format_run(r1, size_pt=9.5, color_rgb=(51, 65, 85))

    doc.add_page_break()

    # ==========================================
    # 2. 13天行程總覽速查表 (Overview)
    # ==========================================
    p_ov_title = doc.add_paragraph()
    p_ov_title.paragraph_format.space_before = Pt(10)
    p_ov_title.paragraph_format.space_after = Pt(12)
    r_ov_t = p_ov_title.add_run("📋 14 天行程大綱與夜宿速查表")
    format_run(r_ov_t, size_pt=16, bold=True, color_rgb=(30, 58, 138))

    ov_table = doc.add_table(rows=1, cols=4)
    ov_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    ov_table.autofit = False
    
    headers = ["天數", "日期 / 樞紐站", "當日主題核心精華", "夜宿地點 / 旅館類型"]
    widths = [Inches(0.8), Inches(1.5), Inches(2.6), Inches(1.6)]

    hdr_cells = ov_table.rows[0].cells
    for i, h_text in enumerate(headers):
        hdr_cells[i].width = widths[i]
        set_cell_background(hdr_cells[i], "1E3A8A")
        set_cell_margins(hdr_cells[i], top=120, bottom=120, left=140, right=140)
        set_cell_border(hdr_cells[i], bottom={'sz': 8, 'val': 'single', 'color': '1E3A8A'})
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h_text)
        format_run(r, size_pt=9.5, bold=True, color_rgb=(255, 255, 255))

    for day in itinerary:
        d_num = day.get('dayNumber', '')
        d_date = day.get('date', '')
        d_base = day.get('baseStation', '').split(' / ')[0]
        d_title = day.get('title', '').replace(f"Day {d_num}：", "")
        d_lodge = day.get('lodging', '')

        row = ov_table.add_row()
        cells = row.cells
        
        bg_col = "F8FAFC" if d_num % 2 == 1 else "FFFFFF"
        for i in range(4):
            cells[i].width = widths[i]
            set_cell_background(cells[i], bg_col)
            set_cell_margins(cells[i], top=100, bottom=100, left=120, right=120)
            set_cell_border(cells[i], bottom={'sz': 4, 'val': 'single', 'color': 'E2E8F0'},
                                      top={'sz': 4, 'val': 'single', 'color': 'E2E8F0'},
                                      left={'sz': 4, 'val': 'single', 'color': 'E2E8F0'},
                                      right={'sz': 4, 'val': 'single', 'color': 'E2E8F0'})
        
        # 天數
        p0 = cells[0].paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r0 = p0.add_run(f"D{d_num}")
        format_run(r0, size_pt=9.5, bold=True, color_rgb=(37, 99, 235))

        # 日期
        p1 = cells[1].paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r1 = p1.add_run(f"{format_word_date_with_dow(d_date)}\n({d_base})")
        format_run(r1, size_pt=8.5, color_rgb=(71, 85, 105))

        # 標題
        p2 = cells[2].paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r2 = p2.add_run(d_title)
        format_run(r2, size_pt=9, bold=True, color_rgb=(30, 58, 138))

        # 住宿
        p3 = cells[3].paragraphs[0]
        p3.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r3 = p3.add_run(d_lodge)
        format_run(r3, size_pt=8.5, color_rgb=(51, 65, 85))

    doc.add_page_break()

    # ==========================================
    # 3. 每日獨立篇章 (每天自成一個段落，強制分頁)
    # ==========================================
    for day_idx, day in enumerate(itinerary):
        d_num = day.get('dayNumber', day_idx + 1)
        d_date = day.get('date', '')
        d_base = day.get('baseStation', '')
        d_lodging = day.get('lodging', '')
        d_loc = day.get('lodgingLocation', '')
        d_recom = day.get('recommendedHotels', '')
        d_notes = day.get('notes', '')
        spot_ids = day.get('spotIds', [])

        # 1. 每日醒目主標題（簡短化，全景點列入）
        short_title = get_short_day_title(day, spots_map)
        p_day_title = doc.add_paragraph()
        p_day_title.paragraph_format.space_before = Pt(8)
        p_day_title.paragraph_format.space_after = Pt(3)
        r_day_t = p_day_title.add_run(f"{short_title}")
        format_run(r_day_t, size_pt=13.5, bold=True, color_rgb=(30, 58, 138))

        # 2. 單一行精簡摘要列（取代原本分割的基本資訊表格）
        pass_status = "🎫 JR Pass 7日券【適用日】(全額免費搭乘指定席)" if 3 <= d_num <= 9 else "💳 刷交通 IC 卡 / 現金 (非 Pass 天數)"
        p_meta = doc.add_paragraph()
        p_meta.paragraph_format.space_before = Pt(0)
        p_meta.paragraph_format.space_after = Pt(6)
        r_meta = p_meta.add_run(f"🗓️ {format_word_date_with_dow(d_date)} ｜ 樞紐站：{d_base} ｜ 票券：{pass_status} ｜ 🏨 宿：{d_lodging} ({d_loc})")
        format_run(r_meta, size_pt=8.5, bold=True, color_rgb=(71, 85, 105))

        # 3. 當日全日整合行程流程表 (從頭看到尾的一整天流程)
        p_tbl_hdr = doc.add_paragraph()
        p_tbl_hdr.paragraph_format.space_before = Pt(4)
        p_tbl_hdr.paragraph_format.space_after = Pt(3)
        r_tbh = p_tbl_hdr.add_run("📋 當日全日整合行程全流程表 (動線・景點・蓋章・美食・交通與住宿)")
        format_run(r_tbh, size_pt=10.5, bold=True, color_rgb=(37, 99, 235))

        # 取得全日各流程節點
        day_schedule = build_full_day_schedule(d_num, itinerary, spots_map, fares_list)

        master_table = doc.add_table(rows=1, cols=5)
        master_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        master_table.autofit = False
        m_widths = [Inches(0.8), Inches(1.7), Inches(0.8), Inches(1.0), Inches(2.2)]
        m_headers = ["時段 / 階段", "行程項目 / 動線", "類別 / 工具", "費用 / 票券", "詳細備忘・蓋章處・參訪重點與美食叮嚀"]

        for hi, htext in enumerate(m_headers):
            hc = master_table.rows[0].cells[hi]
            hc.width = m_widths[hi]
            set_cell_background(hc, "1E293B")
            set_cell_margins(hc, top=60, bottom=60, left=60, right=60)
            hp = hc.paragraphs[0]
            hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            hr = hp.add_run(htext)
            format_run(hr, size_pt=8, bold=True, color_rgb=(255, 255, 255))

        for r_idx, step in enumerate(day_schedule, 1):
            s_time, s_item, s_mode, s_cost, s_details = step
            row = master_table.add_row()
            rcells = row.cells
            bg = "F8FAFC" if r_idx % 2 == 1 else "FFFFFF"

            for ci in range(5):
                rcells[ci].width = m_widths[ci]
                set_cell_background(rcells[ci], bg)
                set_cell_margins(rcells[ci], top=50, bottom=50, left=60, right=60)
                set_cell_border(rcells[ci], bottom={'sz': 4, 'val': 'single', 'color': 'E2E8F0'},
                                            top={'sz': 4, 'val': 'single', 'color': 'E2E8F0'},
                                            left={'sz': 4, 'val': 'single', 'color': 'E2E8F0'},
                                            right={'sz': 4, 'val': 'single', 'color': 'E2E8F0'})

            # 0: 時段
            p0 = rcells[0].paragraphs[0]
            p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r0 = p0.add_run(s_time)
            format_run(r0, size_pt=8, bold=True, color_rgb=(30, 58, 138))

            # 1: 行程項目
            p1 = rcells[1].paragraphs[0]
            r1 = p1.add_run(s_item)
            format_run(r1, size_pt=8.5, bold=True, color_rgb=(15, 23, 42))

            # 2: 類別 / 工具
            p2 = rcells[2].paragraphs[0]
            p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r2 = p2.add_run(s_mode)
            format_run(r2, size_pt=8, color_rgb=(71, 85, 105))

            # 3: 費用 / 票券
            p3 = rcells[3].paragraphs[0]
            p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
            is_pass = "Pass免費" in s_cost or "Pass 全額" in s_cost
            cost_color = (16, 185, 129) if is_pass else ((220, 38, 38) if "自費" in s_cost else (71, 85, 105))
            r3 = p3.add_run(s_cost)
            format_run(r3, size_pt=7.5, bold=True, color_rgb=cost_color)

            # 4: 詳細備忘與叮嚀
            p4 = rcells[4].paragraphs[0]
            r4 = p4.add_run(s_details)
            format_run(r4, size_pt=7.5, color_rgb=(51, 65, 85))

        # 4. 時刻表附件與憑證圖檔（緊接在全流程表下方，不分段）
        att_dict = day.get('transitAttachments', {})
        valid_items = []
        seen_keys = set()
        seen_day_imgs = set()

        alias_groups = [
            {'tomonoura_onsen_to_bitchu_matsuyama', f'morning_origin_{d_num}', 'fukuyama_to_bitchu_matsuyama'},
            {'okayama_hotel_to_okayama_airport', f'morning_origin_{d_num}', f'okayama_airport_{d_num}', f'tpe_flight_{d_num}'},
            {'shin_yamaguchi_lodging_to_kintaikyo', f'morning_origin_{d_num}'},
            {'hagi_lodging_to_motonosumi_jinja', f'morning_origin_{d_num}'},
            {'taikodani_inari_to_kaike_onsen', 'tsuwano_to_kaike_onsen'},
            {'yasugi_lodging_to_gassan_toda', 'yonago_hotel_to_gassan_toda'},
            {'hiroshima_lodging_d9_to_yoshida_koriyama', f'morning_origin_{d_num}'}
        ]

        def mark_seen(key):
            seen_keys.add(key)
            for grp in alias_groups:
                if key in grp:
                    seen_keys.update(grp)

        prev_day_lodging_id = None
        for p_idx in range(day_idx - 1, -1, -1):
            prev_sids = itinerary[p_idx].get('spotIds', [])
            if prev_sids:
                prev_day_lodging_id = prev_sids[-1]
                break

        active_spot_set = set(spot_ids)
        if prev_day_lodging_id:
            active_spot_set.add(prev_day_lodging_id)

        for s_i, sid in enumerate(spot_ids):
            prev_id = spot_ids[s_i - 1] if s_i > 0 else prev_day_lodging_id
            candidate_keys = []
            if s_i == 0 and sid == 'okayama_airport':
                if d_num == 14:
                    candidate_keys.extend(["okayama_hotel_to_okayama_airport", f"morning_origin_{d_num}", f"okayama_airport_{d_num}"])
                else:
                    candidate_keys.extend([f"tpe_flight_{d_num}", f"okayama_airport_{d_num}"])
            if prev_id:
                candidate_keys.append(f"{prev_id}_to_{sid}")
            if s_i == 0:
                if d_num in (12, 13):
                    candidate_keys.extend(["tomonoura_onsen_to_bitchu_matsuyama", "fukuyama_to_bitchu_matsuyama"])
                candidate_keys.extend([f"morning_origin_{d_num}", f"start_to_{sid}"])

            for ck in candidate_keys:
                if ck in att_dict and ck not in seen_keys:
                    item = att_dict[ck]
                    if (item.get('images') and len(item['images']) > 0) or item.get('notes') or item.get('link'):
                        mark_seen(ck)
                        dedup_imgs = []
                        for im in item.get('images', []):
                            u = im.get('url', '')
                            if u and u not in seen_day_imgs:
                                seen_day_imgs.add(u)
                                dedup_imgs.append(im)
                        if dedup_imgs or item.get('notes') or item.get('link'):
                            if d_num in (12, 13) and ck in {'tomonoura_onsen_to_bitchu_matsuyama', f'morning_origin_{d_num}', 'fukuyama_to_bitchu_matsuyama'}:
                                title = "鞆之浦溫泉 ➔ 福山 ➔ 備中松山城 (晨間出發接駁)"
                            elif d_num == 14:
                                title = "岡山站西口 ➔ 岡山桃太郎機場 (機場利木津巴士直達・2027/01/25 IT215)"
                            else:
                                title = get_segment_display_title(ck, day, spots_map)
                            new_item = dict(item)
                            new_item['images'] = dedup_imgs
                            valid_items.append((ck, title, new_item))

        for k, item in att_dict.items():
            if k not in seen_keys:
                belongs = False
                if '_to_' in k:
                    f_sid, t_sid = k.split('_to_', 1)
                    belongs = (f_sid in active_spot_set and t_sid in active_spot_set)
                elif k.endswith(f'_{d_num}'):
                    belongs = True
                if not belongs:
                    continue
                if (item.get('images') and len(item['images']) > 0) or item.get('notes') or item.get('link'):
                    mark_seen(k)
                    dedup_imgs = []
                    for im in item.get('images', []):
                        u = im.get('url', '')
                        if u and u not in seen_day_imgs:
                            seen_day_imgs.add(u)
                            dedup_imgs.append(im)
                    if dedup_imgs:
                        new_item = dict(item)
                        new_item['images'] = dedup_imgs
                        valid_items.append((k, get_segment_display_title(k, day, spots_map), new_item))

        if valid_items:
            p_att_sec = doc.add_paragraph()
            p_att_sec.paragraph_format.space_before = Pt(8)
            p_att_sec.paragraph_format.space_after = Pt(2)
            r_asc = p_att_sec.add_run("📎 流程對應：乘車時刻表附件與票券圖檔")
            format_run(r_asc, size_pt=9.5, bold=True, color_rgb=(30, 58, 138))

            for v_idx, (k, seg_title, item) in enumerate(valid_items):
                p_seg = doc.add_paragraph()
                p_seg.paragraph_format.space_before = Pt(5)
                p_seg.paragraph_format.space_after = Pt(2)
                p_seg.paragraph_format.left_indent = Inches(0.1)
                r_seg = p_seg.add_run(f"🚅 【搭乘附件】{seg_title}")
                format_run(r_seg, size_pt=9.0, bold=True, color_rgb=(30, 58, 138))

                notes_text = item.get('notes', '').strip()
                if notes_text:
                    p_nt = doc.add_paragraph()
                    p_nt.paragraph_format.space_before = Pt(1)
                    p_nt.paragraph_format.space_after = Pt(2)
                    p_nt.paragraph_format.left_indent = Inches(0.25)
                    r_nt = p_nt.add_run(f"📝 班次備忘：{notes_text}")
                    format_run(r_nt, size_pt=8.0, color_rgb=(51, 65, 85))

                link_text = item.get('link', '').strip()
                if link_text:
                    p_lk = doc.add_paragraph()
                    p_lk.paragraph_format.space_before = Pt(1)
                    p_lk.paragraph_format.space_after = Pt(2)
                    p_lk.paragraph_format.left_indent = Inches(0.25)
                    r_lk = p_lk.add_run(f"🔗 時刻表連結：{link_text}")
                    format_run(r_lk, size_pt=7.5, color_rgb=(37, 99, 235))

                imgs = item.get('images', [])
                if imgs:
                    for img_obj in imgs:
                        img_url = img_obj.get('url', '')
                        img_name = img_obj.get('name', '時刻表縮圖')
                        if img_url:
                            insert_timetable_image(doc, img_url, img_name)

        # 每一天的結尾：如果不是最後一天，強制分頁！
        if day_idx < len(itinerary) - 1:
            doc.add_page_break()

    # 儲存檔案（同時儲存主檔與最新版檔案）
    try:
        doc.save("山陰山陽14天旅行小冊子.docx")
        print("成功產生 Word 小冊子：山陰山陽14天旅行小冊子.docx")
    except Exception as e:
        print(f"主檔鎖定中: {e}")
    try:
        doc.save("山陰山陽14天旅行小冊子_最新版.docx")
        print("同步產生 Word 小冊子：山陰山陽14天旅行小冊子_最新版.docx")
    except Exception as e:
        print(f"最新版鎖定中: {e}")

if __name__ == "__main__":
    create_booklet_word()
