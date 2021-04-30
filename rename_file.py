import os
import xlwt
import re
import importlib
import sys
import time
import io

from pdfminer.pdfparser import PDFParser, PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LTTextBoxHorizontal, LAParams
from pdfminer.pdfinterp import PDFTextExtractionNotAllowed



import win32com,re
import os
import docx;
import xlrd
import xlwt
from win32com.client import Dispatch, constants
from docx import Document;



root=r'E:\校级教师成果奖'
workbook = xlwt.Workbook(encoding = 'ascii');
worksheet = workbook.add_sheet('My Worksheet');
renamesheet = workbook.add_sheet('rename Worksheet');
nameofproject = workbook.add_sheet('nameofproject Worksheet');

sys.stdout = io.TextIOWrapper(sys.stdout.buffer,encoding='utf-8')
importlib.reload(sys)
time1 = time.time()
#for root,dirs,files in os.walk(dir):
style = xlwt.XFStyle() # 初始化样式
style.alignment.wrap = 1#自动换行
k=0;
k1=0;
def iter_dir(dir):
	global k;
	for subdirs in os.listdir(dir):
		if '.' in subdirs:
			col=os.path.join(dir, subdirs).count("\\")
			worksheet.write(k, col-2, os.path.join(dir, subdirs))
			if col == 3:
				name_list = ["申请书", "佐证材料", "申报书", "申报材料",'总结报告']
				flag=1;
				for name in name_list:
					if name in subdirs:
						renamesheet.write(k, col-2, os.path.join(dir, re.findall(r"\d+\.?\d*", dir)[0]+name+os.path.splitext(subdirs)[1]))
						try:
							os.rename(os.path.join(dir, subdirs), os.path.join(dir, re.findall(r"\d+\.?\d*", dir)[0]+name+os.path.splitext(subdirs)[1]))
						except:
							a=1;
						flag=0;
						break
				if flag==1:
					renamesheet.write(k, col - 2, os.path.join(dir, subdirs))
			else:
				renamesheet.write(k, col - 2, os.path.join(dir, subdirs))
			#读取pdf文档
			if col == 3:#判断文件是否为二级目录下的文件
				name_shenbao = ["申请书", "申报书", "申报材料"]#
				for files in name_shenbao:
					if files in subdirs:
						if os.path.splitext(subdirs)[1]=='.pdf':
							parse(os.path.join(dir, re.findall(r"\d+\.?\d*", dir)[0]+files+os.path.splitext(subdirs)[1]))

			k=k+1;
		else:
			col = os.path.join(dir, subdirs).count("\\")
			worksheet.write(k, col-2, os.path.join(dir, subdirs))
			if col==2:
				renamesheet.write(k, col - 2, os.path.join(dir, re.findall(r"\d+\.?\d*", subdirs)[0]))
				os.rename(os.path.join(dir, subdirs),os.path.join(dir, re.findall(r"\d+\.?\d*", subdirs)[0]))
				iter_dir(os.path.join(dir, os.path.join(dir, re.findall(r"\d+\.?\d*", subdirs)[0])))
			else:
				renamesheet.write(k, col - 2, os.path.join(dir, subdirs))
				iter_dir(os.path.join(dir, os.path.join(dir, subdirs)))
			k = k + 1;

	workbook.save('Excel_Workbook.xls')
def parse(text_path):
	global k1;
	#解析PDF文本，并保存到TXT文件中
	fp = open(text_path, 'rb')
	# 用文件对象创建一个PDF文档分析器
	parser = PDFParser(fp)
	# 创建一个PDF文档
	doc = PDFDocument()
	# 连接分析器，与文档对象
	parser.set_document(doc)
	doc.set_parser(parser)
	# 提供初始化密码，如果没有密码，就创建一个空的字符串3++-
	doc.initialize()
	# 检测文档是否提供txt转换，不提供就忽略
	if not doc.is_extractable:
		raise PDFTextExtractionNotAllowed
	else:
		# 创建PDF，资源管理器，来共享资源
		rsrcmgr = PDFResourceManager()
		# 创建一个PDF设备对象
		laparams = LAParams()
		device = PDFPageAggregator(rsrcmgr, laparams=laparams)
		# 创建一个PDF解释其对象
		interpreter = PDFPageInterpreter(rsrcmgr, device)

		# 循环遍历列表，每次处理一个page内容
		# doc.get_pages() 获取page列表
		flag_break=1;
		for page in doc.get_pages():
			if flag_break==0:
				break
			interpreter.process_page(page)
			# 接受该页面的LTPage对象
			layout = device.get_result()
			# 这里layout是一个LTPage对象 里面存放着 这个page解析出的各种对象
			# 一般包括LTTextBox, LTFigure, LTImage, LTTextBoxHorizontal 等等
			# 想要获取文本就获得对象的text属性，
			writeresult='';
			for x in layout:
				if (isinstance(x, LTTextBoxHorizontal))and '成果名称' in x.get_text():
					writeresult=writeresult+x.get_text();
					for x in layout:
						print(x)
				if (isinstance(x, LTTextBoxHorizontal))and '成果主要完成人' in x.get_text():
					flag_break=0;
					break
		nameofproject.write(k1, 0, writeresult);
		k1 = k1 + 1;
iter_dir(root);
a=1;




"""


for sundirs in os.listdir(dir):
    if '.' in sundirs:
        worksheet.write(k, 1, sundirs)
        k =k+1;
    else:
        worksheet.write(k-1, 1, sundirs)
        for subdirs in os.listdir(os.path.join(dir, sundirs)):
            subdirs


    k = k + 1;
    print(dirs)
    #for sunfile in dirs:
       # print(dir, sunfile);
        #print(os.path.join(dir, sunfile))
a=1;
               #worksheet.write(0, 0, label='Row 0, Column 0 Value');
        #workbook.save('Excel_Workbook.xls');
        """

'''Examples Generating Excel Documents Using Python’s xlwt

Here are some simple examples using Python’s xlwt library to dynamically generate Excel documents.

Please note a useful alternative may be ezodf, which allows you to generate ODS (Open Document Spreadsheet) files for LibreOffice / OpenOffice. You can check them out at:http://packages.python.org/ezodf/index.html

The Simplest Example'''

'''workbook = xlwt.Workbook(encoding = 'ascii')
worksheet = workbook.add_sheet('My Worksheet')
worksheet.write(0, 0, label = 'Row 0, Column 0 Value')
workbook.save('Excel_Workbook.xls')

Formatting the Contents of a Cell
import xlwt
workbook = xlwt.Workbook(encoding = 'ascii')
worksheet = workbook.add_sheet('My Worksheet')
font = xlwt.Font() # Create the Font
font.name = 'Times New Roman'
font.bold = True
font.underline = True
font.italic = True
style = xlwt.XFStyle() # Create the Style
style.font = font # Apply the Font to the Style
worksheet.write(0, 0, label = 'Unformatted value')
worksheet.write(1, 0, label = 'Formatted value', style) # Apply the Style to the Cell
workbook.save('Excel_Workbook.xls')

Attributes of the Font Object
font.bold = True # May be: True, False
font.italic = True # May be: True, False
font.struck_out = True # May be: True, False
font.underline = xlwt.Font.UNDERLINE_SINGLE # May be: UNDERLINE_NONE, UNDERLINE_SINGLE, UNDERLINE_SINGLE_ACC, UNDERLINE_DOUBLE, UNDERLINE_DOUBLE_ACC
font.escapement = xlwt.Font.ESCAPEMENT_SUPERSCRIPT # May be: ESCAPEMENT_NONE, ESCAPEMENT_SUPERSCRIPT, ESCAPEMENT_SUBSCRIPT
font.family = xlwt.Font.FAMILY_ROMAN # May be: FAMILY_NONE, FAMILY_ROMAN, FAMILY_SWISS, FAMILY_MODERN, FAMILY_SCRIPT, FAMILY_DECORATIVE
font.charset = xlwt.Font.CHARSET_ANSI_LATIN # May be: CHARSET_ANSI_LATIN, CHARSET_SYS_DEFAULT, CHARSET_SYMBOL, CHARSET_APPLE_ROMAN, CHARSET_ANSI_JAP_SHIFT_JIS, CHARSET_ANSI_KOR_HANGUL, CHARSET_ANSI_KOR_JOHAB, CHARSET_ANSI_CHINESE_GBK, CHARSET_ANSI_CHINESE_BIG5, CHARSET_ANSI_GREEK, CHARSET_ANSI_TURKISH, CHARSET_ANSI_VIETNAMESE, CHARSET_ANSI_HEBREW, CHARSET_ANSI_ARABIC, CHARSET_ANSI_BALTIC, CHARSET_ANSI_CYRILLIC, CHARSET_ANSI_THAI, CHARSET_ANSI_LATIN_II, CHARSET_OEM_LATIN_I
font.colour_index = ?
font.get_biff_record = ?
font.height = 0x00C8 # C8 in Hex (in decimal) = 10 points in height.
font.name = ?
font.outline = ?
font.shadow = ?

Setting the Width of a Cell
import xltw
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('My Sheet')
worksheet.write(0, 0, 'My Cell Contents')
worksheet.col(0).width = 3333 # 3333 = 1" (one inch).
workbook.save('Excel_Workbook.xls')

Entering a Date into a Cell
import xlwt
import datetime
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('My Sheet')
style = xlwt.XFStyle()
style.num_format_str = 'M/D/YY' # Other options: D-MMM-YY, D-MMM, MMM-YY, h:mm, h:mm:ss, h:mm, h:mm:ss, M/D/YY h:mm, mm:ss, [h]:mm:ss, mm:ss.0
worksheet.write(0, 0, datetime.datetime.now(), style)
workbook.save('Excel_Workbook.xls')

Adding a Formula to a Cell
import xlwt
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('My Sheet')
worksheet.write(0, 0, 5) # Outputs 5
worksheet.write(0, 1, 2) # Outputs 2
worksheet.write(1, 0, xlwt.Formula('A1*B1')) # Should output "10" (A1[5] * A2[2])
worksheet.write(1, 1, xlwt.Formula('SUM(A1,B1)')) # Should output "7" (A1[5] + A2[2])
workbook.save('Excel_Workbook.xls')

Adding a Hyperlink to a Cell
import xlwt
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('My Sheet')
worksheet.write(0, 0, xlwt.Formula('HYPERLINK("http://www.google.com";"Google")')) # Outputs the text "Google" linking to http://www.google.com
workbook.save('Excel_Workbook.xls')

Merging Columns and Rows
import xlwt
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('My Sheet')
worksheet.write_merge(0, 0, 0, 3, 'First Merge') # Merges row 0's columns 0 through 3.
font = xlwt.Font() # Create Font
font.bold = True # Set font to Bold
style = xlwt.XFStyle() # Create Style
style.font = font # Add Bold Font to Style
worksheet.write_merge(1, 2, 0, 3, 'Second Merge', style) # Merges row 1 through 2's columns 0 through 3.
workbook.save('Excel_Workbook.xls')

Setting the Alignment for the Contents of a Cell
import xlwt
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('My Sheet')
alignment = xlwt.Alignment() # Create Alignment
alignment.horz = xlwt.Alignment.HORZ_CENTER # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
alignment.vert = xlwt.Alignment.VERT_CENTER # May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
style = xlwt.XFStyle() # Create Style
style.alignment = alignment # Add Alignment to Style
worksheet.write(0, 0, 'Cell Contents', style)
workbook.save('Excel_Workbook.xls')

Adding Borders to a Cell
# Please note: While I was able to find these constants within the source code, on my system (using LibreOffice,) I was only presented with a solid line, varying from thin to thick; no dotted or dashed lines.
import xlwt
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('My Sheet')
borders = xlwt.Borders() # Create Borders
borders.left = xlwt.Borders.DASHED # May be: NO_LINE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUM_DASHED, THIN_DASH_DOTTED, MEDIUM_DASH_DOTTED, THIN_DASH_DOT_DOTTED, MEDIUM_DASH_DOT_DOTTED, SLANTED_MEDIUM_DASH_DOTTED, or 0x00 through 0x0D.
borders.right = xlwt.Borders.DASHED
borders.top = xlwt.Borders.DASHED
borders.bottom = xlwt.Borders.DASHED
borders.left_colour = 0x40
borders.right_colour = 0x40
borders.top_colour = 0x40
borders.bottom_colour = 0x40
style = xlwt.XFStyle() # Create Style
style.borders = borders # Add Borders to Style
worksheet.write(0, 0, 'Cell Contents', style)
workbook.save('Excel_Workbook.xls')

Setting the Background Color of a Cell
import xlwt
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('My Sheet')
pattern = xlwt.Pattern() # Create the Pattern
pattern.pattern = xlwt.Pattern.SOLID_PATTERN # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
pattern.pattern_fore_colour = 5 # May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray, the list goes on...
style = xlwt.XFStyle() # Create the Pattern
style.pattern = pattern # Add Pattern to Style
worksheet.write(0, 0, 'Cell Contents', style)
workbook.save('Excel_Workbook.xls')

TODO: Things Left to Document
- Panes -- separate views which are always in view
- Border Colors (documented above, but not taking effect as it should)
- Border Widths (document above, but not working as expected)
- Protection
- Row Styles
- Zoom / Manification
- WS Props?'''
