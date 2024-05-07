#!/usr/bin/env python3

from datetime import datetime
import pandas as pd
from fpdf import FlexTemplate,FPDF
import os

###############
#  constants  #
###############

image_dir     = "pictures" # the directory for the images
icon_dir      = "icons"    # the directory for icons
default_image = os.path.join(icon_dir, "anonymous.jpg") # path to the default image

########################
#  layout definitions  #
########################

text_elements = [
        { 'name': 'id',               'type': 'T', 'font': 'kaiti', 'align': 'C',  'size': 45, 'x1': 0.5,  'y1': 0.25, 'x2': 3,   'y2': 2 },
        { 'name': 'key',              'type': 'I',    'align': 'C',   'size': 45, 'x1': 0.5,   'y1': 0.25, 'x2': 3,    'y2': 2 },
        { 'name': 'english_name',     'type': 'T',    'font': 'hp', 'align': 'C',   'size': 8, 'x1': 0.5,  'y1': 0.3,  'x2': 2,   'y2': 0.4 },
        { 'name': 'chinese_name',     'type': 'T',    'font': 'hp', 'align': 'C',   'size': 8, 'x1': 2,    'y1': 0.3,  'x2': 3,   'y2': 0.4 },
        { 'name': 'children',         'type': 'T',    'font': 'hp', 'align': 'C',  'size': 10, 'x1': 0.2,  'y1': 0.5,  'x2': 3.3, 'y2': 0.6 },
        { 'name': 'children_chinese', 'type': 'T',    'font': 'hp', 'align': 'C',  'size': 10, 'x1': 0.2,  'y1': 0.5,  'x2': 3.3, 'y2': 0.6 },
        { 'name': 'address',          'type': 'T',    'font': 'hp', 'align': 'C',  'size': 10, 'x1': 0.2,  'y1': 0.5,  'x2': 3.3, 'y2': 0.6 },
        { 'name': 'phone',            'type': 'T',    'font': 'hp', 'align': 'C',  'size': 10, 'x1': 0.2,  'y1': 0.5,  'x2': 3.3, 'y2': 0.6 },
        { 'name': 'email',            'type': 'T',    'font': 'hp', 'align': 'C',  'size': 10, 'x1': 0.2,  'y1': 0.5,  'x2': 3.3, 'y2': 0.6 },
        ]

# read by default 1st sheet of an excel file
df = pd.read_excel('info.xlsx')
fpdf = FPDF(orientation="landscape", format="letter", unit="in")
fpdf.add_font("kaiti", fname="./simkai.ttf")
fpdf.add_font("hp", fname="./HPSimplified.ttf")

#######################
#  utility functions  #
#######################

def is_ascii(s):
    try:
        s.encode().decode('ascii')
    except UnicodeDecodeError:
        return False
    else:
        return True

def is_img(f, d=image_dir, ext=['png','jpg','jpeg']):
    """ take a file without extension, try to detect whether it is an image file """
    return any([ os.path.isfile(os.path.join(d, f) + '.' + e) for e in ext ])

def get_img(f, default=default_image, d=image_dir, ext=['png','jpg','jpeg']):
    """ get the 

    :f: the image file name, with or without extension
    :default: the default image if none exists
    :d: directory for images
    :ext: the possible extensions for images
    :returns: the first image file that exists

    """
    if os.path.splitext(f)[1] == "":
        # no extension
        possible_paths = [ os.path.join(d,f) + '.' + e for e in ext ]
    else:
        possible_paths = [ f ]
    return next(( img for img in possible_paths if os.path.isfile(img) ), default)

def nstr(s):
    return str(s) if pd.notnull(s) else ""

def nint(s):
    try:
        return int(s)
    except ValueError:
        return s

for i, row in df.iterrows():
    # print("{} : {}".format(i, row))
    # print("is_img(row['key']): {}".format(is_img(nstr(row['key']))))
    print("get_img(nstr(row['key'])): {}".format(get_img(nstr(row['key']))))
    t1 = FlexTemplate(fpdf, elements=text_elements)
    fpdf.add_page(orientation="Landscape", format=(1.25, 3.5))
    for e in text_elements:
        print("row[e['name']]: {}".format(row[e['name']]))
        # t1[e['name']] = nstr(row[e['name']])
    # t1.render()
    # print("id", row['id'])
    # for n in range(5):
    #     label=f"name{n}"
    #     name=nstr(row.get(label))
    #     if n == 0 or name != "":
    #         t1 = FlexTemplate(fpdf, elements=elements)
    #         fpdf.add_page(orientation="Landscape", format=(2.25, 3.5))
    #         img = img_map.get(row['church'], img_map['default'])
    #         fpdf.image(img, x=0, y=0, w=3.5, keep_aspect_ratio=True)
    #         if is_ascii(name):
    #             t1['name_EN'] = name
    #         else:
    #             t1['name_CN'] = name
    #         t1['room'] = nstr(row.get('room'))
    #         t1['group'] = nstr(row.get('group'))
    #         print(f"name{i}_{n}: {name}")
    #         t1['id'] = nstr(nint(row['id'])) + (f"_{n}" if n > 0 else "")
    #         t1.render()
    if i >= 5:
        break
fpdf.output(f"out_{datetime.now().strftime('%Y%m%d%H%M%S')}.pdf")
