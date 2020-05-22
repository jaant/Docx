import re

def iter_markdown_spans(doc, state, s):
    span_to_text = dict(t=s.text, br='\n', cNvPr=f'![{s.image}]({s.image})', footnoteReference=f'[^{s.footnote_nr}]')
    head_spaces, body_text, tail_text = re.match(r'(\s*?)(\S.*?|)([\s.,;]*)$', span_to_text[s.xml_tag], re.DOTALL).groups()
    if s.para_id != s.prev.para_id:
        if s.para_style.startswith('Heading'):
            yield '#' * int(s.para_style[7:]) + ' '
        elif s.list_info is not None:
            yield '\n' if s.prev.para_style.startswith('Heading') else ''
            yield ' ' + ('   ' * s.list_info.level)
            yield '1. ' if s.list_info.type == 'decimal' else '* '
        elif s.left_indent:
            yield '> '
    yield head_spaces
    if s.italic and (not s.prev.italic or s.para_id != s.prev.para_id):
        state['italic_markup_symbol'] = '_' if s.bold else '*'
        yield state['italic_markup_symbol']
    if s.bold and (not s.prev.bold or s.para_id != s.prev.para_id):
        yield '**'
    if s.href and s.href != s.prev.href:
        yield '['
    yield body_text
    if s.href and s.next.href != s.href:
        yield f']({s.href})'
    if s.bold and (not s.next.bold or s.next.para_id != s.para_id):
        yield '**'
    if s.italic and (not s.next.italic or s.next.para_id != s.para_id):
        yield state['italic_markup_symbol']
    yield tail_text.rstrip() if s.next.para_id != s.para_id or s.next.xml_tag == 'br' else tail_text
    if s.next.para_id != s.para_id:
        if s.next.list_info is not None or s.next.para_id is None:
            yield '\n'
        elif s.left_indent and s.next.left_indent and not s.next.list_info:
            yield '\n>\n'
        else:
            yield '\n\n\n' if s.next.para_style in ('Heading1', 'Heading2') else '\n\n'

def convert(doc):
    state = {}
    return ''.join(md_span for doc_span in doc.text_spans for md_span in iter_markdown_spans(doc, state, doc_span))
