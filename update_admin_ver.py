import re
import sys
sys.stdout.reconfigure(encoding='utf-8')

# 1. Update admin.html
with open('admin.html', 'r', encoding='utf-8') as f:
    admin_content = f.read()

# Update script tag query param
admin_content = admin_content.replace(
    'src="spots_data.js?v=2027_01_12_ROUTE_14DAYS_YAMAGUCHI_V87"',
    'src="spots_data.js?v=2027_01_12_ROUTE_14DAYS_YAMAGUCHI_V91_D5_KAIKE_D6_TSUWANO"'
)

# Update CURRENT_ITINERARY_VER
new_ver = "2027_01_12_ROUTE_14DAYS_YAMAGUCHI_V91_D5_KAIKE_D6_TSUWANO"
admin_content = re.sub(
    r'const CURRENT_ITINERARY_VER\s*=\s*"[^"]+";',
    f'const CURRENT_ITINERARY_VER = "{new_ver}";',
    admin_content
)

# Update getEffectiveDefaultItinerary to support window.EXTERNAL_SPOTS_DATA
old_fn = """function getEffectiveDefaultItinerary() {
      if (typeof window !== 'undefined' && window.INITIAL_SPOTS_DATA && Array.isArray(window.INITIAL_SPOTS_DATA.itinerary) && window.INITIAL_SPOTS_DATA.itinerary.length > 0) {
        return window.INITIAL_SPOTS_DATA.itinerary;
      }
      if (typeof DEFAULT_ITINERARY !== 'undefined' && Array.isArray(DEFAULT_ITINERARY) && DEFAULT_ITINERARY.length > 0) {
        return DEFAULT_ITINERARY;
      }
      return BUILTIN_DEFAULT_ITINERARY;
    }"""

new_fn = """function getEffectiveDefaultItinerary() {
      if (typeof window !== 'undefined' && window.INITIAL_SPOTS_DATA && Array.isArray(window.INITIAL_SPOTS_DATA.itinerary) && window.INITIAL_SPOTS_DATA.itinerary.length > 0) {
        return window.INITIAL_SPOTS_DATA.itinerary;
      }
      if (typeof window !== 'undefined' && window.EXTERNAL_SPOTS_DATA && Array.isArray(window.EXTERNAL_SPOTS_DATA.itinerary) && window.EXTERNAL_SPOTS_DATA.itinerary.length > 0) {
        return window.EXTERNAL_SPOTS_DATA.itinerary;
      }
      if (typeof DEFAULT_ITINERARY !== 'undefined' && Array.isArray(DEFAULT_ITINERARY) && DEFAULT_ITINERARY.length > 0) {
        return DEFAULT_ITINERARY;
      }
      return BUILTIN_DEFAULT_ITINERARY;
    }"""

if old_fn in admin_content:
    admin_content = admin_content.replace(old_fn, new_fn)
    print("Updated getEffectiveDefaultItinerary in admin.html")
else:
    print("Warning: old_fn not found in admin.html")

# Update upgrade logic to skip overwriting D4, D5, D6, D7 from old cache
admin_content = admin_content.replace(
    'if (day.dayNumber === 4) return; // D4 依使用者指定升級至最新車次與丹後周遊券，不以舊快取覆蓋',
    'if ([4, 5, 6, 7].includes(day.dayNumber)) return; // D4-D7 依使用者指定順暢排法升級至最新皆生與津和野行程，不以舊快取覆蓋'
)

# Update notes mentions in admin.html spots
admin_content = admin_content.replace(
    'D5 結束津和野行程後搭特急直達米子，轉公車 15 分進駐皆生溫泉，換穿浴衣享受湧自日本海海底的鹽化物化妝水神湯與海鮮晚宴；隔日清晨自皆生出發攻登月山富田城交通順暢！',
    'D5 結束安來與松江城行程後，搭特急直達米子轉公車進駐皆生溫泉，換穿浴衣享受日本海海底鹽化物化妝水神湯與松葉蟹晚宴；隔日清晨出發直達出雲大社！'
)

with open('admin.html', 'w', encoding='utf-8') as f:
    f.write(admin_content)
print("Saved admin.html successfully.")

# 2. Update spots_data.js to also export window.INITIAL_SPOTS_DATA and window.DEFAULT_ITINERARY
with open('spots_data.js', 'r', encoding='utf-8') as f:
    js_txt = f.read()

prefix = """window.EXTERNAL_SPOTS_DATA = """
if js_txt.startswith(prefix):
    # Add extra globals at end
    suffix_code = """
if (typeof window !== 'undefined') {
  window.INITIAL_SPOTS_DATA = window.EXTERNAL_SPOTS_DATA;
  window.DEFAULT_ITINERARY = window.EXTERNAL_SPOTS_DATA.itinerary;
  window.DEFAULT_SPOTS = window.EXTERNAL_SPOTS_DATA.spots;
}
"""
    if "window.INITIAL_SPOTS_DATA =" not in js_txt:
        js_txt = js_txt.rstrip() + "\n" + suffix_code
        with open('spots_data.js', 'w', encoding='utf-8') as f:
            f.write(js_txt)
        print("Updated spots_data.js with compatibility globals.")
