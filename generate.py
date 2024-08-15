#!/usr/bin/env python3

# TODO:
# Script:
# - Extract the configs to yaml
# Documentation:
# - How to run and set up on Windows?
# - Example excel sheet and output.
# - Demo for the options.
# <2024-06-16, David Deng> #

from datetime import datetime
import pandas as pd
from fpdf import FlexTemplate,FPDF,TitleStyle
from fpdf.outline import OutlineSection
from fpdf.enums import XPos
import imagesize
import os
import sys

###########
#  Paths  #
###########

image_dir     = "pictures" # the directory for the images
icon_dir      = "icons"    # the directory for icons
default_image = os.path.join(icon_dir, "anonymous.jpg") # path to the default image
# path to the icons
icons         = {
        "email":    os.path.join(icon_dir, "email.png"),
        "children": os.path.join(icon_dir, "children.png"),
        "address":  os.path.join(icon_dir, "address.png"),
        "phone":    os.path.join(icon_dir, "phone.png"),
}

############
#  Styles  #
############

gen_toc          = True # Generate table of content?
gen_page_num     = True # Generate page number?
gen_header       = True # Generate header indicating section?
symmetric_layout = True # Symmetric Layout, where the image and description will be reversed on even pages
reverse_layout   = True # Reverse the order of image and description on all pages

###############
#  Constants  #
###############

num_per_page = 3 # number of items per page

# All units are in inches

info_width = 2
img_width   = 2 # image width, including margin
img_margin  = 0.1 # margin around the image
item_height = 2.0
header_height = 0.2
footer_height = 0.2

item_scale  = 1

page_height = item_height * num_per_page + \
        (header_height if gen_header else 0) + \
        (footer_height if gen_page_num else 0)
page_width  = info_width + img_width

########################
#  layout definitions  #
########################

elements = [
        { 'name': 'english_name',     'type': 'T', 'font': 'hp',    'multiline': True, 'align': 'L', 'size': 9, 'x1': 0.2,  'x2': 1.8, 'y1': 0.1,  'y2': 0.2 },
        { 'name': 'chinese_name',     'type': 'T', 'font': 'kaiti', 'multiline': True, 'align': 'L', 'size': 7, 'x1': 0.2,  'x2': 1.8, 'y1': 0.40, 'y2': 0.45 },
        { 'name': 'children',         'type': 'T', 'font': 'hp',    'multiline': True, 'align': 'L', 'size': 7, 'x1': 0.2,  'x2': 1.8, 'y1': 0.55, 'y2': 0.7 },
        { 'name': 'children_icon',    'type': 'I', 'font': None,    'multiline': None, 'align': 'L', 'size': 7, 'x1': 0.05, 'x2': 0.2, 'y1': 0.55, 'y2': 0.7 },
        { 'name': 'children_chinese', 'type': 'T', 'font': 'kaiti', 'multiline': True, 'align': 'L', 'size': 7, 'x1': 0.2,  'x2': 1.8, 'y1': 0.70, 'y2': 0.9 },
        { 'name': 'address',          'type': 'T', 'font': 'hp',    'multiline': True, 'align': 'L', 'size': 7, 'x1': 0.2,  'x2': 1.8, 'y1': 0.85, 'y2': 1.0 },
        { 'name': 'address_icon',     'type': 'I', 'font': None,    'multiline': None, 'align': 'L', 'size': 7, 'x1': 0.05, 'x2': 0.2, 'y1': 0.85, 'y2': 1.0 },
        { 'name': 'phone',            'type': 'T', 'font': 'hp',    'multiline': True, 'align': 'L', 'size': 7, 'x1': 0.2,  'x2': 1.8, 'y1': 1.15, 'y2': 1.3 },
        { 'name': 'phone_icon',       'type': 'I', 'font': None,    'multiline': None, 'align': 'L', 'size': 7, 'x1': 0.05, 'x2': 0.2, 'y1': 1.15, 'y2': 1.3 },
        { 'name': 'email',            'type': 'T', 'font': 'hp',    'multiline': True, 'align': 'L', 'size': 7, 'x1': 0.2,  'x2': 1.8, 'y1': 1.6, 'y2': 1.75 },
        { 'name': 'email_icon',       'type': 'I', 'font': None,    'multiline': None, 'align': 'L', 'size': 7, 'x1': 0.05, 'x2': 0.2, 'y1': 1.6, 'y2': 1.75 },
        ]

#######################
#  Utility Functions  #
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

def p(pdf, text, **kwargs):
    "Inserts a paragraph"
    pdf.multi_cell(
        w=pdf.epw,
        h=pdf.font_size,
        text=text,
        new_x="LMARGIN",
        new_y="NEXT",
        **kwargs,
    )

##########
#  Main  #
##########

def render_toc(pdf, outline):
    pdf.set_auto_page_break(True, margin=0.3)
    print("len(outline): {}".format(len(outline)))
    pdf.y += 0.1
    pdf.x = pdf.l_margin
    pdf.set_font("Courier", size=7)
    for section in outline:
        link = pdf.add_link()
        pdf.set_link(link, page=section.page_number)
        text = f'{" " * section.level * 2} {section.name} {"." * (48 - section.level*2 - len(section.name))} {section.page_number}'
        print("text: {}".format(text))
        p(pdf, text, align="J", link=link)

def gen_pdf(df, fpdf, title=None):
# read by default 1st sheet of an excel file
    fpdf.set_auto_page_break(False)

    for i, row in df.iterrows():
        img = get_img(nstr(row['key']))
        print("row['english_name']: {}".format(row['english_name']))

        # add a new page if needed, not for the first page, because we already had a page
        if i % num_per_page == 0 and i != 0:
            fpdf.add_page(orientation="Portrait", format=(page_width, page_height))

        # render header
        if i % num_per_page == 0 and gen_header and title is not None:
            with fpdf.local_context():
                old_y = fpdf.y
                fpdf.set_y(0.2 * header_height)
                fpdf.set_font('hp',size=8)
                fpdf.cell(0, header_height, text=title, align='C')
                fpdf.set_y(old_y)

        # render footer
        if i % num_per_page == 0 and gen_page_num:
            with fpdf.local_context():
                old_y = fpdf.y
                fpdf.set_y(-1.2 * footer_height)
                fpdf.set_font('hp',size=8)
                fpdf.cell(0, footer_height, text=f'Page { fpdf.page_no() }', align='C')
                fpdf.set_y(old_y)

        with fpdf.local_context():
            fpdf.set_section_title_styles(
                # Level 0 titles: a hack to render the text on invisible areas
                TitleStyle(
                    font_family="hp",
                    font_size_pt=2,
                    t_margin=0,
                    l_margin=-1000,
                    b_margin=0,
                ),
                # Level 1 titles:
                TitleStyle(
                    font_family="hp",
                    font_size_pt=2,
                    t_margin=0,
                    l_margin=-1000,
                    b_margin=0,
                ),
            )
            if i == 0 and title is not None:
                print("title: {}".format(title))
                fpdf.start_section(title, level=0)
            fpdf.start_section(row['english_name'], level=1)

        even_page = (i // num_per_page) % 2 == 0

        # fill in information
        info = FlexTemplate(fpdf, elements=elements)
        for e in elements:
            key = e['name']
            if key in row:
                info[key] = nstr(row[key])
            elif key.endswith("_icon"):
                pre = key.removesuffix("_icon")
                if pre in row and pd.notnull(row[pre]):
                    info[key] = icons[pre]

        # do some computation on the layout
        reverse = (even_page and symmetric_layout) ^ reverse_layout # whether to reverse image and info or not for this item

        img_x = img_margin if reverse else info_width + img_margin  # img on the
        info_x = img_margin * 2 + img_width if reverse else 0

        img_y = header_height + img_margin + item_height * (i % num_per_page)
        info_y = header_height + (i % num_per_page) * item_height

        w,h = imagesize.get(img)
        if w > h:
            # fit by width
            img_w = item_scale * (img_width - img_margin * 2)
            img_h = 0
        else:
            # fit by height
            img_w = 0
            img_h = item_scale * (item_height - img_margin * 2)

        # render info
        with fpdf.local_context():
            info.render(
                    offsetx=info_x,
                    offsety=info_y,
                    scale=item_scale)

            # render image
            fpdf.image(
                    img,
                    img_x,
                    img_y,
                    img_w,
                    img_h,
                    )


fpdf = FPDF(orientation="portrait", format=(page_width, page_height), unit="in")
fpdf.add_font("kaiti", fname="./simkai.ttf")
fpdf.add_font("hp", fname="./HPSimplified_Rg.ttf")

df_current = pd.read_excel(f"info_new.xlsx", sheet_name="info_current")
df_previous = pd.read_excel(f"info_new.xlsx", sheet_name="info_previous")

fpdf.add_page()
fpdf.set_y(0.2)
fpdf.set_font("hp", size=15)
with fpdf.local_context():
    p(fpdf, "GLCAC Directory List 2024", align="C")

if gen_toc:
    fpdf.insert_toc_placeholder(render_toc, 2)

gen_pdf(df_current, fpdf, "Current Members/Adherence")
fpdf.add_page()
gen_pdf(df_previous, fpdf, "Previous Members/Adherence")

dest = f"out_{datetime.now().strftime('%Y%m%d%H%M%S')}.pdf"

if len(sys.argv) == 2:
    dest = sys.argv[1]

fpdf.output(dest)
