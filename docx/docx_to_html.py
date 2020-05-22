import html

def indented(list_info, text):
    return (list_info.level * '    ' if list_info else '') + text

def iter_html_spans(s, doc=None, prefix=''):
    if s.para_id != s.prev.para_id:
        if s.list_info:
            if not s.prev.list_info or s.list_info.level > s.prev.list_info.level:
                yield indented(s.list_info, '<ol>\n' if s.list_info.type == 'decimal' else '<ul>\n')
            elif s.list_info.level == s.prev.list_info.level:
                yield indented(s.list_info, '</li>\n')
            yield indented(s.list_info, f'<li value="{s.list_info.number}">')
        elif s.left_indent:
            yield '<blockquote>\n  ' if not s.prev.left_indent else '  '
        if s.para_style.startswith('Heading'):
            yield '<h%s>' % int(s.para_style[7:])
        else:
            yield '<p>'
        yield prefix
    if s.italic and (not s.prev.italic or s.para_id != s.prev.para_id):
        yield '<i>'
    if s.bold and (not s.prev.bold or s.para_id != s.prev.para_id):
        yield '<b>'
    if s.href and s.href != s.prev.href:
        yield f'<a href="{s.href}">'
    if s.image:
        yield f'<img src="{s.image}">'
    else:
        yield dict(t=html.escape(s.text), br='<br>\n', cNvPr='[IMAGE]', footnoteReference=f'<sup>{s.footnote_nr}</sup>')[s.xml_tag]
    if s.href and s.next.href != s.href:
        yield '</a>'
    if s.bold and (not s.next.bold or s.next.para_id != s.para_id):
        yield '</b>'
    if s.italic and (not s.next.italic or s.next.para_id != s.para_id):
        yield '</i>'
    if s.next.para_id != s.para_id:
        if s.para_style.startswith('Heading'):
            yield '</h%s>' % int(s.para_style[7:])
        else:
            yield '</p>'
        if doc:
            for footnote_html in generate_html_footnotes(f for f in doc.footnotes if f.para_id == s.para_id):
                yield footnote_html
            for comment_html in generate_html_comments(t for t in doc.comment_threads if t.para_id == s.para_id):
                yield comment_html
        if s.list_info:
            yield '\n'
            if not s.next.list_info or s.next.list_info.level < s.list_info.level:
                i, tgt_lvl = s.list_info, (s.next.list_info.level if s.next.list_info else -1)
                while i and i.level > tgt_lvl:
                    yield indented(i, '</li>\n')
                    yield indented(i, '</ol>\n' if i.type == 'decimal' else '</ul>\n')
                    i = i.parent
                yield indented(i, '</li>\n') if i else ''
        elif s.left_indent:
            yield '\n</blockquote>\n' if not s.next.left_indent or s.next.list_info else '\n'
        else:
            yield '\n'

def generate_html(text_spans, doc=None, prefix=''):
    text_spans_with_prefix = ((text_span, prefix if nr == 0 else '') for nr, text_span in enumerate(text_spans))
    return ''.join(html_span for text_span, prefix in text_spans_with_prefix for html_span in iter_html_spans(text_span, doc, prefix))

def generate_html_footnotes(footnotes):
    for footnote in footnotes:
        yield generate_html(footnote.text_spans, doc=None, prefix=f'<sup>{footnote.nr}</sup>').replace('<p>', '<p class="metapara">')

def generate_html_comments(comment_threads):
    for thread in comment_threads:
        yield '<p class="metapara">\n<i>&gt; %s</i>\n</p>\n' % html.escape(thread.quote)
        for comment in thread.comments:
            comment_prefix = '<b>[%s %s]</b> ' % (comment.datetime[:10], html.escape(comment.author))
            yield generate_html(comment.text_spans, doc=None, prefix=comment_prefix).replace('<p>', '<p class="metapara">')

def convert(doc):
    html_title = f'<title>{html.escape(doc.title)}</title>'
    html_style = '<style>.metapara {margin-left: 60px; margin-right: 60px; background: #f0f0f0;}</style>'
    html_head = ['<head>', '<meta http-equiv="content-type" content="text/html; charset=utf-8" />', html_title, html_style, '</head>']
    html_body = ['<body>', generate_html(doc.text_spans, doc).strip(), '</body>']
    html_lines = ['<!DOCTYPE html>', '<html lang="en">'] + html_head + html_body + ['</html>', '']
    return '\n'.join(html_lines)
