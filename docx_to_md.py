import sys; sys.dont_write_bytecode = True
import argparse, pathlib
import docx

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('docx_file', type=argparse.FileType('rb'), help='docx file to be converted')
    parser.add_argument('--overwrite', action='store_true')
    args = parser.parse_args()
    doc_path = pathlib.Path(args.docx_file.name)
    doc = docx.load_doc(doc_path)
    args.docx_file.close()
    md_path = doc_path.with_suffix('.md')
    if not md_path.exists() or args.overwrite:
        text_md = docx.convert_to_markdown(doc)
        md_path.open('w', encoding='utf8', newline='').write(text_md)
        print(f'{md_path} written')
    else:
        print(f'{md_path} already exists!')
