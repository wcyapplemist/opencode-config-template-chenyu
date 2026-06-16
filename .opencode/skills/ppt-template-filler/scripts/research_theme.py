"""Extract theme colors and fonts from template.pptx."""
import sys
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, '.')
from pptx import Presentation
from pptx.oxml.ns import qn
from lxml import etree
from io import BytesIO

NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'

prs = Presentation('templates/template.pptx')
master = prs.slide_masters[0]

theme_part = master.part.part_related_by(
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme'
)
theme_xml = etree.parse(BytesIO(theme_part.blob)).getroot()

theme_elements = theme_xml.find(f'{{{NS_A}}}themeElements')
clr_scheme = theme_elements.find(f'{{{NS_A}}}clrScheme')
font_scheme = theme_elements.find(f'{{{NS_A}}}fontScheme')

print('=' * 60)
print('THEME COLORS (clrScheme)')
print('=' * 60)
for child in clr_scheme:
    tag = etree.QName(child).localname
    for color_elem in child:
        color_tag = etree.QName(color_elem).localname
        if color_tag == 'srgbClr':
            val = color_elem.get('val')
            r, g, b = int(val[0:2], 16), int(val[2:4], 16), int(val[4:6], 16)
            print(f'  {tag:15s}: #{val}  RGB({r:3d}, {g:3d}, {b:3d})')
        elif color_tag == 'sysClr':
            val = color_elem.get('lastClr', '')
            sys_val = color_elem.get('val', '')
            r, g, b = int(val[0:2], 16), int(val[2:4], 16), int(val[4:6], 16)
            print(f'  {tag:15s}: #{val}  RGB({r:3d}, {g:3d}, {b:3d})  (sysClr={sys_val})')

print()
print('=' * 60)
print('THEME FONTS (fontScheme)')
print('=' * 60)
major = font_scheme.find(f'{{{NS_A}}}majorFont')
minor = font_scheme.find(f'{{{NS_A}}}minorFont')
for label, font_elem in [('Major (heading)', major), ('Minor (body)', minor)]:
    latin = font_elem.find(f'{{{NS_A}}}latin')
    ea = font_elem.find(f'{{{NS_A}}}ea')
    latin_tp = latin.get('typeface') if latin is not None else '?'
    ea_tp = ea.get('typeface') if ea is not None else '?'
    print(f'  {label}: latin="{latin_tp}", ea="{ea_tp}"')
