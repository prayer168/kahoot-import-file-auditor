#!/usr/bin/env python3
"""Read-only Kahoot XLSX extraction. Writes JSON; never modifies a workbook.

Requires openpyxl. Use a currently verified JSON --rules profile for optional
platform limits. Content, curriculum and actual import remain unverified.
"""
import argparse
import hashlib
import io
import json
import math
import re
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path

import openpyxl


FIELDS = ['question', 'answer1', 'answer2', 'answer3', 'answer4', 'time', 'correct']


def text(value):
    if value is None:
        return ''
    if isinstance(value, float) and math.isfinite(value) and value.is_integer():
        return str(int(value))
    return str(value)


def nonblank(value):
    return bool(text(value).strip())


def header_key(value):
    value = re.sub(r'\s+', ' ', text(value).strip().lower())
    if re.match(r'^question(?:\s|$|[-(])', value) or value in ('題目', '題幹'):
        return 'question'
    match = re.match(r'^answer\s*([1-4])(?:\s|$|[-(])', value)
    if match:
        return 'answer' + match[1]
    match = re.fullmatch(r'(?:答案|選項)\s*([1-4])', value)
    if match:
        return 'answer' + match[1]
    if value.startswith('correct answer') or value in ('正確答案', '正確答案編號'):
        return 'correct'
    if value.startswith('time limit') or value in ('時間限制', '作答時間'):
        return 'time'
    return None


def issue(code, level, detail, **location):
    return dict(code=code, level=level, detail=detail, **location)


def parse_answers(value):
    raw = text(value).strip()
    if not re.fullmatch(r'[1-4](?:\s*,\s*[1-4])*', raw):
        return None
    return [int(v.strip()) for v in raw.split(',')]


def check_item(item, rules=None):
    values = item['values']
    formulas = set(item.get('formula_fields', []))
    problems = []
    def add(code, level, detail):
        problems.append(issue(code, level, detail))
    if formulas:
        add('formula_cell', 'pending', '公式未執行；須核對顯示值與匯入行為：' + ', '.join(sorted(formulas)))
    if 'question' not in formulas and not nonblank(values.get('question')):
        add('empty_question', 'A', '有選項或答案資料，但題幹空白。')
    options = [values.get('answer' + str(i)) for i in range(1, 5)]
    if all(k in values for k in ('answer1', 'answer2')) and not formulas.intersection({'answer1', 'answer2', 'answer3', 'answer4'}):
        if sum(nonblank(v) for v in options) < 2:
            add('too_few_options', 'A', '可讀的非空白選項少於兩個。')
        normalized = [text(v).strip() for v in options if nonblank(v)]
        if len(set(normalized)) < len(normalized):
            add('duplicate_options', 'B', '選項文字重複，需審查是否造成多解。')
        nonempty_indices = [i for i, v in enumerate(options) if nonblank(v)]
        if nonempty_indices and nonempty_indices != list(range(len(nonempty_indices))):
            add('option_gap', 'pending', '選項中間或開頭有空缺；核對官方範本及答案位置，勿自動重排。')
    if 'correct' in values and 'correct' not in formulas:
        answers = parse_answers(values.get('correct'))
        item['parsed_correct_indices'] = answers
        if answers is None:
            add('invalid_answer_indices', 'A', '答案須為 1–4 的編號，以半形逗號分隔；不得空白或使用未識別格式。')
        else:
            if len(set(answers)) != len(answers):
                add('duplicate_answer_index', 'B', '正確答案編號重複。')
            for index in set(answers):
                if 'answer' + str(index) not in values:
                    add('answer_column_unmapped', 'pending', f'正確答案 {index} 的選項欄未辨識，不能推定為空白。')
                elif 'answer' + str(index) not in formulas and not nonblank(options[index - 1]):
                    add('answer_points_to_blank', 'A', f'正確答案 {index} 指向空白選項。')
    if rules:
        for key in ['question', 'answer1', 'answer2', 'answer3', 'answer4']:
            if key not in values or key in formulas:
                continue
            limit = rules['question_max'] if key == 'question' else rules['answer_max']
            value = text(values.get(key))
            # Code-point count is explicit. Emoji/combining handling is not assumed.
            if len(value) > limit:
                add('length_limit', 'A', f'{key} 有 {len(value)} 個 Unicode 碼點，超過當次規則 {limit}。')
            if any(ord(c) > 0xffff or 0x300 <= ord(c) <= 0x36f for c in value):
                add('unicode_count', 'pending', f'{key} 含特殊 Unicode；另核對平台字數計算。')
        if 'time' in values and 'time' not in formulas:
            raw = values.get('time')
            try:
                number = float(raw)
            except (ValueError, TypeError):
                number = None
            if isinstance(raw, bool) or number not in rules['time_values']:
                add('invalid_time', 'A', '時間不在所提供當次規則的允許值內。')
    item['issues'] = problems
    return item


def inspect_workbook(workbook, source, rules=None):
    result = dict(source=source, status='extracted', content_review='not_started',
                  platform_rules='provided_profile' if rules else 'not_verified', sheets=[], items=[], issues=[])
    for sheet in workbook.worksheets:
        info = dict(name=sheet.title, state=sheet.sheet_state, headers=[], item_count=0)
        result['sheets'].append(info)
        mapping = None
        ordinal = 0
        seen_questions = {}
        for row_no, row in enumerate(sheet.iter_rows(), 1):
            found = {}
            duplicates = []
            for col, cell in enumerate(row):
                key = header_key(cell.value)
                if key:
                    if key in found:
                        duplicates.append(key)
                    found[key] = col
            if 'question' in found and len(found) >= 2:
                mapping = found
                info['headers'].append(dict(row=row_no, columns={k: row[v].coordinate for k, v in found.items()}))
                missing = [key for key in FIELDS if key not in found]
                if missing:
                    result['issues'].append(issue('missing_columns', 'pending', '欄位未辨識：' + ', '.join(missing), sheet=sheet.title, row=row_no))
                if duplicates:
                    result['issues'].append(issue('duplicate_headers', 'pending', '重複欄名，映射有歧義；須手動核對。', sheet=sheet.title, row=row_no))
                continue
            if mapping is None:
                continue
            values = {key: row[col].value for key, col in mapping.items()}
            if not any(nonblank(v) for v in values.values()):
                continue  # Template numbering alone is not a question.
            ordinal += 1
            identity = json.dumps([source, sheet.title, row_no], ensure_ascii=False)
            item = dict(id=hashlib.sha256(identity.encode()).hexdigest()[:20], source=source,
                        sheet=sheet.title, excel_row=row_no, ordinal=ordinal,
                        values=values, cells={k: row[v].coordinate for k, v in mapping.items()},
                        formula_fields=[k for k, v in mapping.items() if row[v].data_type == 'f'],
                        extra_cells={c.coordinate: c.value for i, c in enumerate(row)
                                     if i not in mapping.values() and nonblank(c.value)},
                        review_status='尚未審查')
            check_item(item, rules)
            question = text(values.get('question')).strip()
            if question and question in seen_questions:
                item['issues'].append(issue('repeated_question', 'B', f'題幹與 Excel 第 {seen_questions[question]} 列相同；確認是否重複題。'))
            if question:
                seen_questions[question] = row_no
            result['items'].append(item)
            info['item_count'] += 1
        if not info['headers']:
            info['status'] = 'unmapped'
            result['issues'].append(issue('unmapped_sheet', 'pending', '未辨識題目表頭，須檢視此表內容並決定是否手動映射。', sheet=sheet.title))
        else:
            info['status'] = 'extracted'
    if not result['items']:
        result['issues'].append(issue('no_items', 'pending', '未擷取任何題目，不能宣稱檔案通過。'))
    return result


def inspect_bytes(data, source, rules=None):
    workbook = None
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            media = [name for name in archive.namelist() if name.startswith('xl/media/')]
        workbook = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=False, keep_links=False)
        result = inspect_workbook(workbook, source, rules)
        result['sha256'] = hashlib.sha256(data).hexdigest()
        result['embedded_media'] = media
        if media:
            result['issues'].append(issue('media_review', 'pending', '內含圖片或媒體；本程式不解析，需另行檢視題目依賴的素材。'))
        return result
    except Exception as exc:
        return dict(source=source, status='read_error', content_review='not_started',
                    items=[], issues=[issue('read_error', 'pending', f'{type(exc).__name__}: {exc}')])
    finally:
        if workbook is not None:
            workbook.close()


def iter_inputs(paths):
    seen = set()
    for raw in paths:
        path = Path(raw).resolve()
        candidates = sorted(p for p in path.rglob('*') if p.suffix.lower() in ('.xlsx', '.zip')) if path.is_dir() else [path]
        if not candidates:
            yield str(path), None, '資料夾中無 .xlsx 或 .zip。'
        for file in candidates:
            if file in seen:
                continue
            seen.add(file)
            if file.name.startswith('~$'):
                yield str(file), None, 'Excel 暫存鎖定檔，未當作題庫讀取。'
                continue
            try:
                if file.suffix.lower() == '.zip':
                    with zipfile.ZipFile(file) as archive:
                        members = [m for m in archive.infolist() if m.filename.lower().endswith('.xlsx') and not m.is_dir()]
                        if not members:
                            yield str(file), None, 'ZIP 內沒有 .xlsx；不遞迴解開巢狀 ZIP。'
                        for member in members:
                            # Read directly: never extract paths from an archive onto disk.
                            source = f'{file}!{member.filename}@{member.header_offset}'
                            try:
                                yield source, archive.read(member), None
                            except Exception as exc:
                                yield source, None, str(exc)
                elif file.suffix.lower() == '.xlsx':
                    yield str(file), file.read_bytes(), None
                else:
                    yield str(file), None, '不支援此輸入類型；需要 .xlsx、資料夾或 ZIP。'
            except Exception as exc:
                yield str(file), None, str(exc)


def load_rules(path):
    if path is None:
        return None
    rules = json.loads(Path(path).read_text(encoding='utf-8-sig'))
    for key in ('source_url', 'checked_at', 'question_max', 'answer_max', 'time_values'):
        if not rules.get(key):
            raise ValueError('規格 JSON 缺少 ' + key)
    for key in ('question_max', 'answer_max'):
        if type(rules[key]) is not int or rules[key] <= 0:
            raise ValueError(key + ' 必須是正整數')
    if not isinstance(rules['time_values'], list) or not all(type(v) in (int, float) and v > 0 for v in rules['time_values']):
        raise ValueError('time_values 必須是正數陣列')
    return rules


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('inputs', nargs='+')
    parser.add_argument('--output', required=True, help='New JSON output path; will not overwrite existing files.')
    parser.add_argument('--rules', help='Currently verified platform rules JSON; omitted means limits unchecked.')
    args = parser.parse_args()
    output = Path(args.output).resolve()
    if output.suffix.lower() != '.json' or output.exists():
        parser.error('--output 必須是尚不存在的 .json 路徑。')
    try:
        rules = load_rules(args.rules)
    except (OSError, ValueError) as exc:
        parser.error(str(exc))
    results = []
    for source, data, error in iter_inputs(args.inputs):
        results.append(inspect_bytes(data, source, rules) if error is None else
                       dict(source=source, status='read_error', content_review='not_started', items=[],
                            issues=[issue('input_error', 'pending', error)]))
    report = dict(schema_version=1, generated_at=datetime.now(timezone.utc).isoformat(),
                  rules=rules, content_review='not_started', actual_import='not_tested', files=results,
                  item_count=sum(len(r['items']) for r in results))
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open('x', encoding='utf-8') as stream:
        json.dump(report, stream, ensure_ascii=False, indent=2, default=str, allow_nan=False)
    print(json.dumps(dict(output=str(output), files=len(results), items=report['item_count'],
                          read_errors=sum(r['status'] == 'read_error' for r in results)), ensure_ascii=False))
    return 2 if any(r['status'] == 'read_error' for r in results) else 0


if __name__ == '__main__':
    sys.exit(main())
