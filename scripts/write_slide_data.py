import json
import re
import sys
import os
from pptx import Presentation

if len(sys.argv) < 3:
    raise Exception('Required arguments are: path to pptx and path to a folder with slide images')

pptx_path = sys.argv[1]
images_path = sys.argv[2]

prs = Presentation(pptx_path)

titles = []
max_length = 100

for index, slide in enumerate(prs.slides):
    title = "Untitled " + str(index)
    title_shape = slide.shapes.title
    if title_shape:
        title = re.sub(r"\s+", " ", title_shape.text)
        if len(title) > max_length:
            # TODO: may be better to find the closest space
            title = title[0:max_length]

    titles.append(title)

folder_content = os.listdir(images_path)

slides_data = []

for entry in folder_content:
    if os.path.isdir(os.path.join(images_path, entry)):
        continue

    (_, ext) = os.path.splitext(entry)
    # Names are supposed to be like Slide-0.jpg, Slide-1.jpg etc.
    index_match = re.findall(r"-(\d+)" + re.escape(ext) + r"$", entry)
    if len(index_match) == 0:
        # When there is a single slide there is no index in title
        index = 0
    else:
        index = int(index_match[0])

    entry_title = titles[index]

    slides_data.append({
        "index": index,
        "title": entry_title,
        "image": entry
    })


def sort_by_index(slide):
    return slide["index"]


slides_data.sort(key=sort_by_index)
data_file = None
try:
    data_file = open(os.path.join(images_path, 'data.json'), "w")
    data_file.write(json.dumps({"slides": slides_data}))
finally:
    if data_file is not None:
        data_file.close()
