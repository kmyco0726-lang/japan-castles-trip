import json
import re

# 1. 載入 trip_schedule_data.json
with open('trip_schedule_data.json', 'r', encoding='utf-8') as f:
    days_data = json.load(f)

# 精準交通 link 映射
link_mapping = {
    5: {
        0: 'https://www.jr-odekake.net/', # 米子 ➔ 安來
        1: 'https://www.city.yasugi.shimane.jp/kurashi/sumai/yellow-bus/', # 安來 ➔ 月山富田城 (黃色巴士)
        3: 'https://www.city.yasugi.shimane.jp/kurashi/sumai/yellow-bus/zikokusaisinn.html', # 月山富田城 ➔ 足立美術館 (黃色巴士)
        5: 'https://www.adachi-museum.or.jp/shuttle-bus', # 足立美術館 ➔ 安來 (接駁車)
        7: 'https://www.jr-odekake.net/', # 安來 ➔ 松江
        8: 'https://matsue-bus.jp/lakeline', # 松江 ➔ 松江城大手前 (レイクライン)
        10: 'https://www.jr-odekake.net/', # 松江 ➔ 米子
        11: 'https://hinomarubus.co.jp/timetable_route/2956/?tab=3#up-bus', # 米子 ➔ 皆生溫泉 (日之丸巴士)
    },
    6: {
        0: 'https://hinomarubus.co.jp/timetable_route/2956/?tab=3#up-bus', # 皆生溫泉 ➔ 米子
        1: 'https://www.jr-odekake.net/', # 米子 ➔ 出雲市 (特急やくも)
        2: 'https://www.ichibata.co.jp/bus/rosen/taisha.html', # 出雲市 ➔ 出雲大社 (一畑巴士)
        5: 'https://www.ichibata.co.jp/bus/rosen/taisha.html', # 出雲大社 ➔ 出雲市 (一畑巴士)
        6: 'https://www.jr-odekake.net/', # 出雲市 ➔ 益田 (特急スーパーおき)
    },
    7: {
        0: 'https://www.jr-odekake.net/', # 益田 ➔ 津和野
        3: 'https://www.jr-odekake.net/', # 津和野 ➔ 益田 ➔ 東萩
        4: 'https://www.city.hagi.lg.jp/uploaded/attachment/28216.pdf', # 東萩 ➔ 萩城跡 (萩循環まぁーるバス)
        6: 'https://www.city.hagi.lg.jp/uploaded/attachment/28216.pdf', # 萩城跡 ➔ 萩の御厨 (萩循環まぁーるバス)
    },
    8: {
        0: 'https://www.jr-odekake.net/', # 東萩 ➔ 長門市 ➔ 長門古市
        1: 'https://nanavi.jp/sightseeing/motonosumijinja/', # 長門古市 ➔ 元乃隅神社 (觀光計程車)
        3: 'https://nanavi.jp/sightseeing/motonosumijinja/', # 元乃隅神社 ➔ 長門古市 (觀光計程車)
        4: 'https://www.jr-odekake.net/', # 長門古市 ➔ 防府
        6: 'https://www.jr-odekake.net/', # 防府 ➔ 新山口
    },
    9: {
        0: 'https://akiyoshidai-kanko.com/access/', # 新山口 ➔ 秋芳洞 (防長巴士/秋吉台觀光交通)
        2: 'https://www.chugoku-jrbus.co.jp/general/yamaguchi/', # 秋芳洞 ➔ 香山公園五重塔 (中國JR巴士秋吉線)
        4: '', # 香山公園 ➔ 山口 (刪除連結)
        6: 'https://www.jr-odekake.net/', # 山口 ➔ 新山口 ➔ 廣島 ➔ 岩國
    },
    10: {
        0: '', # 岩國 ➔ 錦帶橋 (刪除失效連結)
        2: '', # 錦帶橋 ➔ 岩國 (刪除失效連結)
        3: 'https://www.jr-odekake.net/', # 岩國 ➔ 宮島口 (JR山陽本線)
        4: 'https://jr-miyajimaferry.co.jp/', # 宮島口 ➔ 宮島 (JR西日本宮島渡輪)
        8: 'https://www.jr-odekake.net/', # 宮島口 ➔ 廣島 (JR山陽本線)
    },
    11: {
        0: 'https://www.hiroden.co.jp/', # 廣島市內廣電路面電車
    },
    12: {
        0: 'https://www.hiroden.co.jp/bus/route/n-72.html', # 廣島 ➔ 安藝高田 (廣電巴士)
        2: 'https://www.hiroden.co.jp/bus/route/n-72.html', # 安藝高田 ➔ 廣島 (廣電巴士)
        3: 'https://www.jr-odekake.net/', # 廣島 ➔ 福山 (山陽新幹線)
        6: 'https://tomotetsu.co.jp/bus/tomo/', # 福山 ➔ 鞆之浦 (鞆鐵巴士)
    },
    13: {
        0: 'https://tomotetsu.co.jp/cont/wp-content/uploads/2026/09/1.20261001jikoku-tomo.pdf', # 鞆之浦 ➔ 福山
        1: 'https://www.jr-odekake.net/', # 福山 ➔ 倉敷 ➔ 備中高梁 (特急やくも)
        2: 'https://takahasikanko.or.jp/', # 備中高梁 ➔ 鞴峠 (觀光乘合計程車)
        5: 'https://transit.yahoo.co.jp/search/result?from=%E5%82%99%E4%B8%AD%E9%AB%98%E6%A2%81&to=%E5%82%99%E5%89%8D%E4%B8%80%E5%AE%AE',
        7: 'https://www.jr-odekake.net/', # 備前一宮 ➔ 吉備津 (JR吉備線)
        9: 'https://transit.yahoo.co.jp/search/result?from=%E5%90%89%E5%82%99%E6%B4%A5&to=%E5%B2%A1%E5%B1%B1',
    },
    14: {
        2: 'https://www.okayama-airport.org/access/bus', # 岡山 ➔ 岡山機場利木津巴士
    }
}

for d in days_data:
    d_num = d.get('dayNumber')
    if d_num in link_mapping:
        steps = d.get('steps', [])
        for s_idx, target_link in link_mapping[d_num].items():
            if s_idx < len(steps):
                steps[s_idx]['link'] = target_link

# 同步補充 attachments 的官方連結
for d in days_data:
    d_num = d.get('dayNumber')
    atts = d.get('transitAttachments', {})
    if d_num == 5:
        if 'yasugi_lodging_to_gassan_toda' in atts:
            atts['yasugi_lodging_to_gassan_toda']['link'] = 'https://www.city.yasugi.shimane.jp/kurashi/sumai/yellow-bus/'
        if 'yonago_hotel_to_gassan_toda' in atts:
            atts['yonago_hotel_to_gassan_toda']['link'] = 'https://www.city.yasugi.shimane.jp/kurashi/sumai/yellow-bus/'
        if 'adachi_museum_to_nobonita_shokudo' in atts:
            atts['adachi_museum_to_nobonita_shokudo']['link'] = 'https://www.adachi-museum.or.jp/shuttle-bus'
    elif d_num == 6:
        if 'kaike_onsen_to_izumo_taisha' in atts:
            atts['kaike_onsen_to_izumo_taisha']['link'] = 'https://www.ichibata.co.jp/bus/rosen/taisha.html'
        if 'izumo_taisha_to_masuda_lodging' in atts:
            atts['izumo_taisha_to_masuda_lodging']['link'] = 'https://www.jr-odekake.net/'
    elif d_num == 7:
        if 'taikodani_inari_to_hagi' in atts:
            atts['taikodani_inari_to_hagi']['link'] = 'https://www.city.hagi.lg.jp/uploaded/attachment/28216.pdf'
    elif d_num == 10:
        if 'iwakuni_to_itsukushima_jinja' in atts:
            atts['iwakuni_to_itsukushima_jinja']['link'] = 'https://jr-miyajimaferry.co.jp/'
        if 'iwakuni_lodging_to_kintaikyo' in atts:
            atts['iwakuni_lodging_to_kintaikyo']['link'] = ''
    elif d_num == 12:
        if 'fukuyama_to_tomonoura_onsen' in atts:
            atts['fukuyama_to_tomonoura_onsen']['link'] = 'https://tomotetsu.co.jp/bus/tomo/'
    elif d_num == 14:
        if 'okayama_hotel_to_okayama_airport' in atts:
            atts['okayama_hotel_to_okayama_airport']['link'] = 'https://www.okayama-airport.org/access/bus'

# 寫入 trip_schedule_data.json
with open('trip_schedule_data.json', 'w', encoding='utf-8') as f:
    json.dump(days_data, f, ensure_ascii=False, indent=2)
print("1. Successfully updated trip_schedule_data.json")

# 寫入 spots_data.json
with open('spots_data.json', 'r', encoding='utf-8') as f:
    spots_json = json.load(f)
spots_json['itinerary'] = days_data
with open('spots_data.json', 'w', encoding='utf-8') as f:
    json.dump(spots_json, f, ensure_ascii=False, indent=2)
print("2. Successfully updated spots_data.json")

# 寫入 spots_data.js
with open('spots_data.js', 'r', encoding='utf-8') as f:
    js_text = f.read()
it_start = js_text.find('"itinerary": [')
if it_start != -1:
    it_end = js_text.rfind(']')
    # 重新構建 spots_data.js
    spots_part = js_text[:it_start]
    it_str = '"itinerary": ' + json.dumps(days_data, ensure_ascii=False, indent=2) + '\n};\n'
    new_js = spots_part + it_str
    with open('spots_data.js', 'w', encoding='utf-8') as f:
        f.write(new_js)
    print("3. Successfully updated spots_data.js")
else:
    print("Warning: itinerary not found in spots_data.js")

# 準備更新 index.html 與 trip.html 中的 DAYS_DATA
days_data_js = json.dumps(days_data, ensure_ascii=False, indent=2)

def update_html_days_data(filename):
    with open(filename, 'r', encoding='utf-8') as f:
        content = f.read()
    
    start_token = 'const DAYS_DATA = ['
    end_token = '];\n'
    idx_start = content.find(start_token)
    if idx_start == -1:
        print(f"Error: {start_token} not found in {filename}")
        return
    
    idx_end = content.find(end_token, idx_start)
    if idx_end == -1:
        print(f"Error: {end_token} not found after DAYS_DATA in {filename}")
        return
    
    new_content = content[:idx_start] + 'const DAYS_DATA = ' + days_data_js + ';' + content[idx_end + len(end_token)-1:]
    
    with open(filename, 'w', encoding='utf-8') as f:
        f.write(new_content)
    print(f"4. Successfully updated DAYS_DATA in {filename}")

update_html_days_data('index.html')
update_html_days_data('trip.html')

