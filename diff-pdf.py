import sys
import os
import argparse
from io import StringIO
from pdfminer.converter import TextConverter, PDFConverter, PDFPageAggregator
from pdfminer.layout import LTTextBox, LAParams
from pdfminer.layout import LTTextLine
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser


class PdfToText(object):
    ROBOT_LIBRARY_SCOPE = 'Global'

    def __init__(self):
        self.codec = 'utf-8'
        self.scale = 1

    def convert_pdf_to_txt(self, path, page_no=-1):
        rsrcmgr = PDFResourceManager()

        retstr = StringIO()
        laparams = LAParams()
        device = TextConverter(rsrcmgr, retstr, laparams=laparams)
        fp = open(path, 'rb')
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        password = ""
        maxpages = 0
        caching = True
        pagenos = set()
        i = 0
        for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password, caching=caching,
                                      check_extractable=True):
            i += 1
            if i != page_no:
                continue
            interpreter.process_page(page)

        fp.close()
        device.close()
        str = retstr.getvalue()
        retstr.close()

        return str

    def read_pdf(self, file_name):
        rsrcmgr = PDFResourceManager()
        laparams = LAParams()
        fp = open(file_name, 'rb')
        parser = PDFParser(fp)
        document = PDFDocument(parser)
        device = PDFPageAggregator(rsrcmgr, laparams=laparams)
        interpreter = PDFPageInterpreter(rsrcmgr, device)

        pages = {}
        for page in PDFPage.create_pages(document):
            dictionary = {}
            dictionary['textbox'] = []
            dictionary['textline'] = []

            interpreter.process_page(page)
            layout = device.get_result()
            for item in layout:
                if isinstance(item, LTTextBox):
                    dictionary['textbox'].append(item)
                    for child in item:
                        if isinstance(child, LTTextLine):
                            dictionary['textline'].append(child)
            pages[layout.pageid] = dictionary
        return pages

    def compare_pdf(self, file1, file2, header_text, x_margin=10, compare_margin=0.2):
        rsrcmgr = PDFResourceManager()
        retstr = StringIO()
        laparams = LAParams()
        fp = open(file1, 'rb')
        parser = PDFParser(fp)
        document = PDFDocument(parser)
        device = PDFPageAggregator(rsrcmgr, laparams=laparams)

        out = StringIO()

        layoutmode = 'normal'
        scale = 1.3
        fontscale = 1

        html_coverter = HTMLPrivateConverter(rsrcmgr, out, scale=scale,
                                             layoutmode=layoutmode, laparams=laparams, fontscale=fontscale,
                                             imagewriter=None, header_text=header_text, x_margin=x_margin)
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        testpages = PDFPage.create_pages(document)

        file_dict = self.read_pdf(file2)

        for page in testpages:
            interpreter.process_page(page)
            layout = device.get_result()

            html_coverter.page_begin(layout)
            if file_dict.get(layout.pageid) == None:
                break
            compare_page = file_dict[layout.pageid]

            for item in layout:
                if isinstance(item, LTTextBox):
                    html_coverter.begin_div('textbox', 1, item.x0 + html_coverter.x_margin, item.y1, item.width,
                                            item.height,
                                            item.get_writing_mode())

                    for child in item:
                        if isinstance(child, LTTextLine):
                            self.compare_textline(child, compare_page, html_coverter, compare_margin)
                            html_coverter.put_newline()
                    html_coverter.end_div()
            html_coverter.page_end()

        fp.close()
        device.close()
        retstr.close()

        return out.getvalue()

    def compare_textline(self, child, compare_page, html_coverter, compare_margin):
        comp_result = True
        for comp_textline in compare_page['textline']:
            if child.x0 - compare_margin < comp_textline.x0 and comp_textline.x0 < child.x0 + compare_margin and child.y1 - compare_margin < comp_textline.y1 and comp_textline.y1 < child.y1 + compare_margin:
                if child.get_text() != comp_textline.get_text():
                    html_coverter.put_text_invalid(child.get_text(), child._objs[0].fontname, child._objs[0].size)
                else:
                    html_coverter.put_text(child.get_text(), child._objs[0].fontname, child._objs[0].size)

                comp_result = False
                break

        if comp_result:
            html_coverter.put_text_invalid(child.get_text(), child._objs[0].fontname, child._objs[0].size)

    def convert_list(self, obj):
        list = []
        for child in obj:
            list.append(child)
        return list


##  HTMLConverter
class HTMLPrivateConverter(PDFConverter):
    RECT_COLORS = {
        'figure': 'yellow',
        'textline': 'magenta',
        'textbox': 'cyan',
        'textgroup': 'red',
        'curve': 'black',
        'page': 'gray',
    }

    TEXT_COLORS = {
        'textbox': 'blue',
        'char': 'black',
    }

    def __init__(self, rsrcmgr, outfp, pageno=1, laparams=None,
                 scale=1, fontscale=1.0, layoutmode='normal', showpageno=True,
                 pagemargin=50, imagewriter=None, header_text='', x_margin=0,
                 rect_colors={'curve': 'black', 'page': 'gray'},
                 text_colors={'char': 'black'}):
        PDFConverter.__init__(self, rsrcmgr, outfp, pageno=pageno, laparams=laparams)
        self.scale = scale
        self.fontscale = fontscale
        self.layoutmode = layoutmode
        self.showpageno = showpageno
        self.pagemargin = pagemargin
        self.imagewriter = imagewriter
        self.rect_colors = rect_colors
        self.text_colors = text_colors
        self.debug = False
        if self.debug:
            self.rect_colors.update(self.RECT_COLORS)
            self.text_colors.update(self.TEXT_COLORS)
        self._yoffset = self.pagemargin
        self._font = None
        self._ffont = ('AllAndNone', 11)
        self._fontstack = []
        self.header_text = header_text
        self.x_margin = x_margin
        self.write_header()
        return

    def write(self, text):
        self.outfp.write(text)
        return

    def write_header(self):
        if self.x_margin != 10:
            self.write('<div><span style="position:absolute; color:%s; left:%dpx; top:%dpx; font-size:%dpx;">' %
                       (1, self.x_margin + 200, 0, 30 * self.scale * 1))
        else:
            self.write('<div><span style="position:absolute; color:%s; left:%dpx; top:%dpx; font-size:%dpx;">' %
                       (1, self.x_margin + 7, 0, 30 * self.scale * 1))
        self.write_text(self.header_text)
        self.write('</span></div>\n')
        return

    def write_footer(self):
        self.write('</body></html>\n')
        return

    def write_text(self, text):
        self.write(text)
        return

    def place_rect(self, color, borderwidth, x, y, w, h):
        color = self.rect_colors.get(color)
        if color is not None:
            self.write('<span style="position:absolute; border: %s %dpx solid; '
                       'left:%dpx; top:%dpx; width:%dpx; height:%dpx;"></span>\n' %
                       (color, borderwidth,
                        x * self.scale, (self._yoffset - y) * self.scale,
                        w * self.scale, h * self.scale))
        return

    def place_border(self, color, borderwidth, item):
        self.place_rect(color, borderwidth, item.x0 + self.x_margin, item.y1, item.width, item.height)
        return

    def place_image(self, item, borderwidth, x, y, w, h):
        if self.imagewriter is not None:
            name = self.imagewriter.export_image(item)
            self.write('<img src="%s" border="%d" style="position:absolute; left:%dpx; top:%dpx;" '
                       'width="%d" height="%d" />\n' %
                       (name, borderwidth,
                        x * self.scale, (self._yoffset - y) * self.scale,
                        w * self.scale, h * self.scale))
        return

    def place_text(self, color, text, x, y, size):
        color = self.text_colors.get(color)
        if color is not None:
            self.write('<span style="position:absolute; color:%s; left:%dpx; top:%dpx; font-size:%dpx;">' %
                       (color, x * self.scale, (self._yoffset - y) * self.scale, size * self.scale * self.fontscale))
            self.write_text(text)
            self.write('</span>\n')
        return

    def begin_div(self, color, borderwidth, x, y, w, h, writing_mode=False):
        self._fontstack.append(self._font)
        self._font = None
        self.write('<div style="position:absolute; border: %s %dpx solid; writing-mode:%s; '
                   'left:%dpx; top:%dpx; width:%dpx; height:%dpx;">' %
                   (color, borderwidth, writing_mode,
                    x * self.scale, (self._yoffset - y) * self.scale,
                    w * self.scale, h * self.scale))
        return

    def end_div(self):
        if self._font is not None:
            self.write('</span>')
        self._font = self._fontstack.pop()
        self.write('</div>')
        return

    def put_text(self, text, fontname, size):
        self.write('<span style="font-family: %s; font-size:%dpx; letter-spacing:-1.2;">' %
                   (fontname, size * self.scale * self.fontscale))
        self.write_text(text)
        self.write('</span>')

        return

    def put_text_invalid(self, text, fontname, size):
        self.write(
            '<span style="font-family: %s; font-size:%dpx; background-color:rgb(255,153,153); letter-spacing:-1.2;">' %
            (fontname, size * self.scale * self.fontscale))
        self.write_text(text)
        self.write('</span>')
        return

    def page_begin(self, layout):
        self._yoffset += layout.y1
        self.place_border('page', 1, layout)
        if self.x_margin != 10:
            self.write('<div style="position:absolute; left:%dpx; top:%dpx;">' % (
                self.x_margin + 200, (self._yoffset - layout.y1) * self.scale))
        else:
            self.write('<div style="position:absolute; left:%dpx; top:%dpx;">' % (
                self.x_margin + 7, (self._yoffset - layout.y1) * self.scale))
        self.write('<a name="%s">Page %s</a></div>\n' % (layout.pageid, layout.pageid))

    def page_end(self):
        self._yoffset += self.pagemargin

    def put_newline(self):
        self.write('<br>\n')
        return

    def close(self):
        self.write_footer()
        return


def merge_html(output_html1, output_html2):
    html = StringIO()
    html.write('<!DOCTYPE HTML PUBLIC>\n')
    html.write('<html><head>\n')
    html.write('<meta http-equiv="Content-Type" content="text/html; charset=%s">\n' % 'utf-8')
    html.write('</head><body>\n')

    html.write('<table><tr><td>\n')
    html.write(output_html1)
    html.write('</td>\n')
    html.write('<td>\n')
    html.write(output_html2)
    html.write('</td></tr></table>\n')

    html.write('</body></html>\n')
    return html.getvalue()


def compare_pdf(file1, file2, output_file, compare_margin=0.2):
    pdf_reader = PdfToText()

    output_html1 = pdf_reader.compare_pdf(file1, file2, 'AS-IS', compare_margin=compare_margin)
    output_html2 = pdf_reader.compare_pdf(file2, file1, 'TO-BE', 650, compare_margin=compare_margin)

    html = merge_html(output_html1, output_html2)

    f_out = open(output_file, 'w', -1, 'utf-8')
    f_out.write(html)


def replace_string_to_sql_format(sql_str):
    sql_str = str(sql_str)
    sql_str = sql_str.replace('\\', '\\\\')
    sql_str = sql_str.replace("'", "\\'")
    return sql_str


def main():
    description = ('Compare PDF documents using PDF Miner '
                   'and print out the differences as HTML documents')
    parser = argparse.ArgumentParser(description=description)

    parser.add_argument('files', nargs='*', help='compare files')
    parser.add_argument('-m', '--compare-margin', default=0.2, type=float, help='object comparison margin')
    parser.add_argument('-o', '--output-file', default='output.html', type=str, help='output file name')

    args = parser.parse_args()

    def error_message(msg):
        sys.stderr.write('ERROR: %s%s' % (msg, os.linesep))
        parser.print_usage(sys.stderr)
        sys.exit(1)

    if len(args.files) != 2:
        error_message('please input two pdf files, ex) python diff-pdf.py <file1> <file2>')
    else:
        compare_pdf(args.files[0], args.files[1], args.output_file, args.compare_margin)


if __name__ == "__main__":
    main()
