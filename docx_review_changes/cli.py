"""
Extract all tracked changes and comments from a .docx file.
Requires Python 3.6+ standard library only.

Usage:
    docx-review-changes yourfile.docx
"""

import sys
import zipfile
import xml.etree.ElementTree as ET

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


def q(name):
    return f'{{{W}}}{name}'


def el_text(el):
    """All w:t text inside el."""
    return ''.join(c.text or '' for c in el.iter(q('t')))


def tokenise_para(para):
    """
    Walk one paragraph and return an ordered list of tokens:
        ('text', str)
        ('del', author, str)
        ('ins', author, str)
        ('comment_start', id)
        ('comment_end', id)
    """
    tokens = []

    def walk(node):
        for child in node:
            tag = child.tag
            if tag == q('commentRangeStart'):
                tokens.append(('comment_start', child.get(q('id'), '')))
            elif tag == q('commentRangeEnd'):
                tokens.append(('comment_end', child.get(q('id'), '')))
            elif tag == q('del'):
                author = child.get(q('author'), 'unknown')
                dt = ''.join(c.text or '' for c in child.iter(q('delText')))
                if dt:
                    tokens.append(('del', author, dt))
            elif tag == q('ins'):
                author = child.get(q('author'), 'unknown')
                it = ''.join(c.text or '' for c in child.iter(q('t')))
                if it:
                    tokens.append(('ins', author, it))
            elif tag == q('r'):
                t = ''.join(c.text or '' for c in child.iter(q('t')))
                if t:
                    tokens.append(('text', t))
            else:
                walk(child)

    walk(para)
    return tokens


def context_window(tokens, index, window=50):
    """
    Return (before, after) plain-text snippets around the token at index.
    """
    before_parts = []
    for i in range(index - 1, -1, -1):
        tok = tokens[i]
        if tok[0] == 'text':
            before_parts.insert(0, tok[1])
        elif tok[0] in ('del', 'ins'):
            before_parts.insert(0, tok[2])
        if sum(len(s) for s in before_parts) >= window:
            break
    before = ''.join(before_parts)[-window:]
    if len(before) == window:
        before = '…' + before

    after_parts = []
    for i in range(index + 1, len(tokens)):
        tok = tokens[i]
        if tok[0] == 'text':
            after_parts.append(tok[1])
        elif tok[0] in ('del', 'ins'):
            after_parts.append(tok[2])
        if sum(len(s) for s in after_parts) >= window:
            break
    after = ''.join(after_parts)[:window]
    if len(after) == window:
        after = after + '…'

    return before, after


def process_para(para, para_num, comment_map):
    tokens = tokenise_para(para)
    results = []
    open_comments = {}  # id -> token index where range started
    i = 0

    while i < len(tokens):
        tok = tokens[i]

        if tok[0] == 'comment_start':
            open_comments[tok[1]] = i

        elif tok[0] == 'comment_end':
            cid = tok[1]
            if cid in open_comments:
                start = open_comments.pop(cid)
                anchored = []
                for j in range(start + 1, i):
                    t = tokens[j]
                    if t[0] == 'text':
                        anchored.append(t[1])
                    elif t[0] in ('del', 'ins'):
                        anchored.append(t[2])
                anchored_text = ''.join(anchored).strip()
                before, after = context_window(tokens, start)
                if cid in comment_map:
                    author, ctext = comment_map[cid]
                    results.append({
                        'type': 'COMMENT',
                        'para': para_num,
                        'author': author,
                        'anchored': anchored_text or '(whole paragraph)',
                        'comment': ctext,
                        'before': before,
                        'after': after,
                    })

        elif tok[0] == 'del':
            author, deleted = tok[1], tok[2]
            before, after = context_window(tokens, i)
            # Look ahead: if the next change token is an ins from the same
            # author, treat del+ins as a single substitution.
            j = i + 1
            while j < len(tokens) and tokens[j][0] not in ('del', 'ins', 'text'):
                j += 1
            if j < len(tokens) and tokens[j][0] == 'ins' and tokens[j][1] == author:
                inserted = tokens[j][2]
                results.append({
                    'type': 'CHANGE',
                    'para': para_num,
                    'author': author,
                    'deleted': deleted,
                    'inserted': inserted,
                    'before': before,
                    'after': after,
                })
                i = j  # skip the paired ins
            else:
                results.append({
                    'type': 'DELETION',
                    'para': para_num,
                    'author': author,
                    'deleted': deleted,
                    'before': before,
                    'after': after,
                })

        elif tok[0] == 'ins':
            author, inserted = tok[1], tok[2]
            before, after = context_window(tokens, i)
            results.append({
                'type': 'INSERTION',
                'para': para_num,
                'author': author,
                'inserted': inserted,
                'before': before,
                'after': after,
            })

        i += 1

    return results


def format_results(all_results):
    if not all_results:
        print("No tracked changes or comments found.")
        return

    for n, r in enumerate(all_results, 1):
        t = r['type']
        print(f"{'─' * 60}")
        print(f"[{n}]  {t}  —  paragraph {r['para']}  —  {r['author']}")
        print()

        if t == 'COMMENT':
            print(f"  Highlighted : \"{r['anchored']}\"")
            print(f"  Comment     : \"{r['comment']}\"")
            print(f"  Context     : {r['before']}[^^^]{r['after']}")

        elif t == 'CHANGE':
            print(f"  Delete  : \"{r['deleted']}\"")
            print(f"  Insert  : \"{r['inserted']}\"")
            print(f"  Context : {r['before']}[^^^]{r['after']}")

        elif t == 'DELETION':
            print(f"  Delete  : \"{r['deleted']}\"")
            print(f"  Context : {r['before']}[^^^]{r['after']}")

        elif t == 'INSERTION':
            print(f"  Insert  : \"{r['inserted']}\"")
            print(f"  Context : {r['before']}[^^^]{r['after']}")

        print()

    print(f"{'─' * 60}")
    print(f"Total: {len(all_results)} item(s)")


def main():
    if len(sys.argv) < 2:
        print("Usage: docx-review-changes yourfile.docx")
        sys.exit(1)

    path = sys.argv[1]

    with zipfile.ZipFile(path, 'r') as z:
        names = z.namelist()
        doc_xml = z.read('word/document.xml')
        cmt_xml = z.read('word/comments.xml') if 'word/comments.xml' in names else None

    comment_map = {}
    if cmt_xml:
        ct = ET.fromstring(cmt_xml)
        for c in ct.findall(f'.//{q("comment")}'):
            cid = c.get(q('id'), '')
            author = c.get(q('author'), 'unknown')
            text = el_text(c).strip()
            comment_map[cid] = (author, text)

    doc = ET.fromstring(doc_xml)
    body = doc.find(f'.//{q("body")}')

    all_results = []
    para_num = 0
    for elem in body:
        if elem.tag == q('p'):
            para_num += 1
            results = process_para(elem, para_num, comment_map)
            all_results.extend(results)

    format_results(all_results)


if __name__ == '__main__':
    main()
