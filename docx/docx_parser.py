import collections, itertools, xml.etree.ElementTree, pathlib, zipfile

### prop definitons ######################################################################################################################

DOC_PROP_DEFS = dict(para_id = {'p.$eid': int},
                     para_style = {'p.pPr.pStyle.val': str, 'default': ''},
                     left_indent = {'p.pPr.ind.left': float, 'default': 0.0},
                     list_level = {'p.pPr.numPr.ilvl.val': int},
                     list_info = {'p.pPr.numPr.numId.val': int},
                     bold = {'r.rPr.b.val': lambda v: bool(int(v)), 'r.rPr.b.$true': bool, 'default': False},
                     italic = {'r.rPr.i.val': lambda v: bool(int(v)), 'r.rPr.i.$true': bool, 'default': False},
                     href = {'hyperlink.id': str},
                     text = {'t.text': str, 'default': ''},
                     image = {'nvPicPr.cNvPr.name': str},
                     footnote_nr = {'footnoteReference.id': lambda v: int(v) + 1},
                     comment_start = {'commentRangeStart.id': int},
                     comment_end = {'commentRangeEnd.id': int})

COMMENT_PROP_DEFS = dict(id = {'comment.id': int},
                         author = {'comment.author': str},
                         datetime = {'comment.date': str},
                         **DOC_PROP_DEFS)

LIST_PROP_DEFS = dict(list_id = {'abstractNum.abstractNumId': int},
                      list_level = {'abstractNum.lvl.ilvl': int},
                      list_start = {'abstractNum.lvl.start.val': int, 'default': 1},
                      list_type = {'abstractNum.lvl.numFmt.val': str},
                      num_id = {'num.numId': int},
                      num_list_id = {'num.abstractNumId.val': int})

REL_PROP_DEFS = dict(href_id = {'Relationship.Id': str},
                     url = {'Relationship.Target': str},
                     type = {'Relationship.TargetMode': str})

### xml parser ###########################################################################################################################

def parse_xml(doc_path, xml_filename, event_types=('start', 'end')):
    with zipfile.ZipFile(doc_path) as zip_file:
        strip_ns = lambda text: text.rsplit('}', 1)[-1] if text.startswith('{') else text
        mk_attrs = lambda e: {strip_ns(k): v for k, v in list(e.attrib.items()) + ([] if e.text is None else [('text', e.text)])}
        filename = f'word/{xml_filename}'
        xml_events = xml.etree.ElementTree.iterparse(zip_file.open(filename), event_types) if filename in zip_file.namelist() else []
        return ((etype, strip_ns(e.tag), mk_attrs(e)) for etype, e in xml_events)

def get_xml_attr_samples(doc_path, xml_filename):
    attrs, scope = {}, []
    for xml_event, xml_tag, xml_attrs in parse_xml(doc_path, xml_filename):
        if xml_event == 'start':
            scope.append(xml_tag)
            new_attrs = {tuple(scope + ([k] if k else [])): v for k, v in list(xml_attrs.items()) + [(None, None)]}
            attrs.update((k, v) for k, v in new_attrs.items() if k not in attrs)
        else:
            scope.pop()
    return {'.'.join(k): attrs[k] for k in sorted(attrs)}

### iter doc spans #######################################################################################################################

def mk_default_span(prop_defs):
    DocSpan = collections.namedtuple('DocSpan', ' '.join(['xml_tag'] + list(prop_defs) + ['prev', 'next']))
    return DocSpan(xml_tag=None, prev=None, next=None, **{k: v.get('default') for k, v in prop_defs.items()})

def mk_prop_extractor(path, prop_name, value_fn):
    scope, attr_name = path[:-1], path[-1]
    def extract_prop(xml_event_nr, xml_attrs, depth=len(scope), attr_name=attr_name, prop_name=prop_name, value_fn=value_fn):
        xml_value = True if attr_name == '$true' else xml_event_nr if attr_name == '$eid' else xml_attrs.get(attr_name)
        return None if xml_value is None else (depth, (prop_name, value_fn(xml_value)))
    return extract_prop

def build_selector_graph_node(graph_root, steps_so_far, prop_defs):
    cur_depth = len(steps_so_far)
    all_paths = {tuple(k.split('.')): (pname, fn) for pname, pdef in prop_defs.items() for k, fn in pdef.items() if k != 'default'}
    matching_paths = [path for path in all_paths if path[:cur_depth] == tuple(steps_so_far)]
    completed_paths = [path for path in matching_paths if path[:-1] == tuple(steps_so_far)]
    extractors = [mk_prop_extractor(path, *all_paths[path]) for path in reversed(completed_paths)] # pdefs overwrite ones to their right
    branch_tags = [path[cur_depth] for path in matching_paths if path not in completed_paths]
    child_nodes = {tag: build_selector_graph_node(graph_root, steps_so_far + [tag], prop_defs) for tag in branch_tags}
    return (child_nodes, extractors)

def do_prop_updates(scope, extractors, xml_event_nr, xml_attrs):
    prop_updates = filter(None, (e(xml_event_nr, xml_attrs) for e in extractors))
    _ = [scope[d][1].append(prop) for depth, prop in prop_updates for d in range(-1, -(depth + 1), -1)]

def iter_doc_spans(doc_filename, xml_filename, prop_defs, trigger_xml_tags):
    default_span = mk_default_span(prop_defs)
    selector_graph_root = build_selector_graph_node({}, [], prop_defs)
    xml_element_stream = enumerate(parse_xml(pathlib.Path(doc_filename).resolve(), xml_filename))
    scope = [(selector_graph_root, [])]
    for xml_event_nr, (xml_event, xml_tag, xml_attrs) in xml_element_stream:
        if xml_event == 'start':
            selector_graph_node = scope[-1][0][0].get(xml_tag) or selector_graph_root[0].get(xml_tag) or ({}, [])
            scope.append((selector_graph_node, []))
            do_prop_updates(scope, selector_graph_node[1], xml_event_nr, xml_attrs)
        else: # xml_event == 'end'
            if 'text' in xml_attrs: # text attr is not always set at the start event
                do_prop_updates(scope, scope[-1][0][1], xml_event_nr, dict(text=xml_attrs['text']))
            if xml_tag in trigger_xml_tags:
                yield default_span._replace(xml_tag=xml_tag, **dict(p for _, props in scope[1:] for p in props))
            scope.pop()

def filter_and_assign_next_prev(filter_fn, doc_spans, prop_defs):
    default_span = mk_default_span(prop_defs)
    filtered_spans = [default_span] + [s for s in doc_spans if filter_fn(s)] + [default_span]
    return [cur._replace(prev=prev, next=next) for prev, cur, next in zip(filtered_spans, filtered_spans[1:], filtered_spans[2:])]

### doc parser ###########################################################################################################################

def populate_list_info(text_spans, list_templates):
    ListInfo = collections.namedtuple('ListInfo', 'parent list_id level type number')
    state = dict(numbers={}, prev_span={})
    def mk_list_info(span):
        if span.list_info:
            prev_span = state['prev_span']
            is_new_para = not prev_span or span.para_id != prev_span.para_id
            key = (span.list_info, span.list_level)
            templ = list_templates[key]
            state['numbers'][key] = state['numbers'].get(key, templ.list_start - 1) + is_new_para
            parent = prev_span.list_info if prev_span else None
            while parent and parent.level >= span.list_level:
                parent = parent.parent
            span = span._replace(list_info=ListInfo(parent, templ.list_id, templ.list_level, templ.list_type, state['numbers'][key]))
        state['prev_span'] = span
        return span
    return (mk_list_info(s) for s in text_spans)

def extract_list_templates(src_path):
    list_spans = list(iter_doc_spans(src_path, 'numbering.xml', LIST_PROP_DEFS, ('lvl', 'num')))
    get_list_spans = lambda list_id: (s for s in list_spans if s.list_id == list_id)
    return {(n.num_id, s.list_level): s for n in list_spans if n.xml_tag == 'num' for s in get_list_spans(n.num_list_id)}

def extract_hrefs(src_path):
    xml_files = ('_rels/document.xml.rels', '_rels/footnotes.xml.rels')
    href_spans = (span for xml_file in xml_files for span in iter_doc_spans(src_path, xml_file, REL_PROP_DEFS, ('Relationship', )))
    return {s.href_id: s.url for s in href_spans}

def extract_footnotes(src_path, hrefs, doc_spans, list_templates):
    set_href = lambda s: s._replace(href=hrefs[s.href]) if s.href else s
    footnote_prop_defs = dict(id={'footnote.id': lambda v: int(v) + 1}, **DOC_PROP_DEFS)
    text_spans = (set_href(s) for s in iter_doc_spans(src_path, 'footnotes.xml', footnote_prop_defs, ('t', 'br')))
    text_spans = populate_list_info(text_spans, list_templates)
    para_ids = {s.footnote_nr: s.para_id for s in doc_spans if s.footnote_nr is not None}
    Footnote = collections.namedtuple('Footnote', 'para_id nr text_spans')
    finalise_spans = lambda spans: filter_and_assign_next_prev(lambda _: True, spans, footnote_prop_defs)
    return sorted(Footnote(para_ids[k], k, finalise_spans(g)) for k, g in itertools.groupby(text_spans, key=lambda s: s.id))

def extract_comment_threads(src_path, hrefs, doc_spans):
    cstarts = {s.comment_start: idx for idx, s in enumerate(doc_spans) if s.xml_tag == 'commentRangeStart'}
    cranges = {s.comment_end: (cstarts[s.comment_end], idx) for idx, s in enumerate(doc_spans) if s.xml_tag == 'commentRangeEnd'}
    assert all(i0 < i1 for i0, i1 in cranges.values())
    if set(cstarts) > set(cranges):
        cranges.update({k: (cstarts[k], cstarts[k]) for k in set(cstarts) - set(cranges)})
    refs = {k: (doc_spans[i1].para_id, tuple(i for i in range(i0, i1) if doc_spans[i].xml_tag == 't')) for k, (i0, i1) in cranges.items()}
    assert set(cstarts) == set(cranges) == set(refs)
    Comment = collections.namedtuple('Comment', 'nr author datetime text_spans')
    set_href = lambda s: s._replace(href=hrefs[s.href]) if s.href else s
    text_spans = [set_href(s) for s in iter_doc_spans(src_path, 'comments.xml', COMMENT_PROP_DEFS, ('t', 'br'))]
    text_span_groups = [(k, list(g)) for k, g in itertools.groupby(sorted(text_spans, key=lambda s: s.id), lambda s: s.id)]
    finalise_spans = lambda spans: filter_and_assign_next_prev(lambda _: True, spans, COMMENT_PROP_DEFS)
    comments = [Comment(nr, ts[0].author, ts[0].datetime, finalise_spans(ts)) for nr, ts in text_span_groups]
    CommentThread = collections.namedtuple('CommentThread', 'para_id quote comments')
    shorten = lambda words: ' '.join(words[:10] + ['/.../'] + words[-10:]) if len(words) > 24 else ' '.join(words)
    quotes = {(para_id, idxs): shorten(''.join(doc_spans[i].text for i in idxs).split(' ')) for (para_id, idxs) in refs.values()}
    default_ref = (max(s.para_id for s in doc_spans), (-1,))
    quotes[default_ref] = '*REFERENCE MISSING*'
    get_ref = lambda c: refs.get(c.nr, default_ref)
    comment_groups = itertools.groupby(sorted(comments, key=lambda c: (get_ref(c), c.datetime)), get_ref)
    return [CommentThread(para_id, quotes[(para_id, idxs)], list(comments)) for (para_id, idxs), comments in comment_groups]

### load #################################################################################################################################

def load(src_path):
    hrefs = extract_hrefs(src_path)
    with zipfile.ZipFile(src_path) as zip_file:
        images = {f.split('/', 2)[2]: zip_file.read(f) for f in zip_file.namelist() if f.startswith('word/media/')}
    doc_xml_tags = ('t', 'br', 'cNvPr', 'footnoteReference', 'commentRangeStart', 'commentRangeEnd')
    doc_spans = list(iter_doc_spans(src_path, 'document.xml', DOC_PROP_DEFS, doc_xml_tags))
    list_templates = extract_list_templates(src_path)
    footnotes = extract_footnotes(src_path, hrefs, doc_spans, list_templates)
    comment_threads = extract_comment_threads(src_path, hrefs, doc_spans)
    text_spans = (s._replace(href=hrefs[s.href]) if s.href else s for s in doc_spans)
    text_spans = populate_list_info(text_spans, list_templates)
    text_spans = filter_and_assign_next_prev(lambda s: not s.xml_tag.startswith('comment'), text_spans, DOC_PROP_DEFS)
    Doc = collections.namedtuple('Doc', 'title text_spans footnotes comment_threads images')
    return Doc(src_path.stem, text_spans, footnotes, comment_threads, images)

### main #################################################################################################################################

if __name__ == '__main__':
    import sys; sys.dont_write_bytecode = True
    import pathlib, docx_to_html, docx_to_markdown
    src_paths = [f for f in pathlib.Path(__file__).parent.glob('*.docx') if f.is_file()]
    for src_path in src_paths:
        print(f'converting {src_path.name}')
        doc = load(src_path)
        src_path.with_suffix('.html').write_text(docx_to_html.convert(doc), encoding='utf8')
        src_path.with_suffix('.md').write_text(docx_to_markdown.convert(doc), encoding='utf8')
