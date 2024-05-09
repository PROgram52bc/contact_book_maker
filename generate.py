#!/usr/bin/env python3

from datetime import datetime
import pandas as pd
# from PIL import Image
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

icons = {
        "email":    os.path.join(icon_dir, "email.png"),
        "children": os.path.join(icon_dir, "children.png"),
        "address":  os.path.join(icon_dir, "address.png"),
        "phone":    os.path.join(icon_dir, "phone.png"),
}

elements = [
        { 'name': 'english_name',     'type': 'T', 'font': 'hp',    'multiline': True, 'align': 'L', 'size': 7, 'x1': 2.2, 'x2': 3.3, 'y1': 0.1,  'y2': 0.2 },
        { 'name': 'chinese_name',     'type': 'T', 'font': 'kaiti', 'multiline': True, 'align': 'L', 'size': 5, 'x1': 2.2, 'x2': 3.3, 'y1': 0.40, 'y2': 0.45 },
        { 'name': 'children',         'type': 'T', 'font': 'hp',    'multiline': True, 'align': 'L', 'size': 5, 'x1': 2.2, 'x2': 3.3, 'y1': 0.55, 'y2': 0.62 },
        { 'name': 'children_icon',    'type': 'I', 'font': None,    'multiline': None, 'align': 'L', 'size': 4, 'x1': 2.15, 'x2': 2.2, 'y1': 0.55, 'y2': 0.62 },
        { 'name': 'children_chinese', 'type': 'T', 'font': 'kaiti', 'multiline': True, 'align': 'L', 'size': 5, 'x1': 2.2, 'x2': 3.3, 'y1': 0.70, 'y2': 0.77 },
        { 'name': 'address',          'type': 'T', 'font': 'hp',    'multiline': True, 'align': 'L', 'size': 4, 'x1': 2.2, 'x2': 3.3, 'y1': 0.85, 'y2': 0.9 },
        { 'name': 'address_icon',     'type': 'I', 'font': None,    'multiline': None, 'align': 'L', 'size': 4, 'x1': 2.15, 'x2': 2.2, 'y1': 0.85, 'y2': 0.9 },
        { 'name': 'phone',            'type': 'T', 'font': 'hp',    'multiline': True, 'align': 'L', 'size': 4, 'x1': 2.2, 'x2': 3.3, 'y1': 1.0,  'y2': 1.05 },
        { 'name': 'phone_icon',       'type': 'I', 'font': None,    'multiline': None, 'align': 'L', 'size': 4, 'x1': 2.15, 'x2': 2.2, 'y1': 1.0,  'y2': 1.05 },
        { 'name': 'email',            'type': 'T', 'font': 'hp',    'multiline': True, 'align': 'L', 'size': 4, 'x1': 2.2, 'x2': 3.3, 'y1': 1.15, 'y2': 1.2 },
        { 'name': 'email_icon',       'type': 'I', 'font': None,    'multiline': None, 'align': 'L', 'size': 4, 'x1': 2.15, 'x2': 2.2, 'y1': 1.15, 'y2': 1.2 },
        ]

# read by default 1st sheet of an excel file
df = pd.read_excel('info.xlsx')
fpdf = FPDF(orientation="landscape", format=(1.6, 3.5), unit="in")
fpdf.set_auto_page_break(False, margin = 0.0)
fpdf.add_font("kaiti", fname="./simkai.ttf")
fpdf.add_font("hp", fname="./HPSimplified_Rg.ttf")

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

def get_img(f, default=default_image, d=image_dir, ext=['png','jpg','jpeg']):
    """ get the image for the given path

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

for i, row in df.iterrows():
    img = get_img(nstr(row['key']))
    print("img: {}".format(img))
    t1 = FlexTemplate(fpdf, elements=elements)
    fpdf.add_page(orientation="Landscape", format=(1.6, 3.5))
    fpdf.image(img, 0.1, 0.1, 0, 1.4)
    for e in elements:
        # print("row[e['name']]: {}".format(row[e['name']]))
        key = e['name']
        if key in row:
            print("key: {}".format(key))
            print("nstr(row[key]): {}".format(nstr(row[key])))
            t1[key] = nstr(row[key])
        elif key.endswith("_icon"):
            pre = key.removesuffix("_icon")
            if pre in row and pd.notnull(row[pre]):
                t1[key] = icons[pre]
    t1.render()
    # if i >= 5:
    #     break
fpdf.output(f"out_{datetime.now().strftime('%Y%m%d%H%M%S')}.pdf")
