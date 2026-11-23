with open('index.html', 'r', encoding='utf-8') as f:
    text = f.read()

import re
pos = text.find('else if (day.dayNumber === 5) {')
pos2 = text.find('else if (day.dayNumber === 6) {', pos)
print(text[pos:pos2])
