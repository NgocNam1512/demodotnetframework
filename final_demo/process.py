import json

path = '01289_lines.json'
datalist = []
with open(path, encoding='utf8') as f:
    data = json.load(f)
    for obj in data:
        new_obj = dict()
        new_obj['label'] = 'line'

        location = dict()
        location['x1'] = min(obj['boundingBox'][0], obj['boundingBox'][2])
        location['y1'] = min(obj['boundingBox'][1], obj['boundingBox'][3])
        location['x2'] = max(obj['boundingBox'][4], obj['boundingBox'][6])
        location['y2'] = max(obj['boundingBox'][5], obj['boundingBox'][7])
        new_obj['location'] = location

        new_obj['content'] = obj['text']
        datalist.append(new_obj)


new_json = dict()
new_json['width'] = 1670
new_json['datalist'] = datalist
with open("01289.json", "w", encoding='utf8') as writer:
    json.dump(new_json, writer, ensure_ascii=False)
