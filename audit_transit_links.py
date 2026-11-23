import json

with open("trip_schedule_data.json", "r", encoding="utf-8") as f:
    trip_data = json.load(f)

for day in trip_data[4:]: # Day 5 to 14
    d_num = day.get("dayNumber")
    title = day.get("title")
    print(f"\n=================== DAY {d_num}: {title} ===================")
    for idx, s in enumerate(day.get("steps", [])):
        m = s.get("mode", "")
        item = s.get("item", "")
        is_transit = any(w in m for w in ['車', '公車', '巴士', '電車', '特急', '新幹線', '纜車', '船', '接駁']) or ('➔' in item)
        if is_transit:
            print(f"Step {idx} [{s.get('time')}]: {item}")
            print(f"   mode: {m}")
            print(f"   link: {s.get('link')}")
            # check if details has urls
            details = s.get('details', '')
            import re
            urls = re.findall(r'https?://[^\s\)\)\"\'，。、<>]+', details)
            print(f"   urls in details: {urls}")
            print(f"   images count: {len(s.get('images', []))}")
            # check attachments
            atts = day.get('transitAttachments', {})
            matched_att_links = []
            for k, att in atts.items():
                if att.get('link'):
                    matched_att_links.append((k, att.get('link')))
            if matched_att_links:
                print(f"   day transitAttachment links: {matched_att_links}")
