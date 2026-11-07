#!/usr/bin/env python3
"""wb_card_pack — генератор наполнения карточек WB из Excel-спецификации.

Вход: xlsx/csv с товарами — колонки артикул/название/бренд/категория +
свободные колонки-характеристики (материал, размер, цвет, габариты...).
Выход: cards_out.xlsx со столбцами:
  артикул | seo_заголовок (<=60 зн) | описание | ключевые_слова | характеристики JSON

Использование:
  python wb_card_pack.py spec.xlsx [--out cards_out.xlsx] [--brand-col бренд]
"""
import argparse, json, re, sys

import openpyxl

STOP = {'и', 'в', 'на', 'с', 'для', 'по', 'из', 'от', 'к', 'у', 'о', 'об',
        'the', 'a', 'of', 'and', 'or', 'to'}


def read_rows(path):
    """xlsx или csv → (headers, rows of dict)"""
    if path.lower().endswith('.csv'):
        import csv
        with open(path, encoding='utf-8-sig', newline='') as f:
            rd = list(csv.reader(f))
        hdr, data = rd[0], rd[1:]
    else:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        it = ws.iter_rows(values_only=True)
        hdr = [str(c or '').strip() for c in next(it)]
        data = list(it)
    rows = []
    for r in data:
        d = {hdr[i]: (str(c).strip() if c is not None else '') for i, c in enumerate(r) if i < len(hdr)}
        if any(d.values()):
            rows.append(d)
    return hdr, rows


def find_col(hdr, names):
    low = {h.lower(): h for h in hdr}
    for n in names:
        if n in low:
            return low[n]
    return None


def seo_title(name, brand, cat):
    """Заголовок <=60 зн: название + бренд (без повторов)."""
    t = name.strip().rstrip('.')
    if brand and brand.lower() not in t.lower():
        t = f'{t} {brand}'
    return t[:60].rstrip()


def keywords(row, cat, brand):
    toks = []
    for v in row.values():
        for w in re.findall(r'[А-Яа-яA-Za-z]{4,}', v):
            lw = w.lower()
            if lw not in STOP and lw not in toks:
                toks.append(lw)
    for w in [cat, brand]:
        if w and w.lower() not in toks:
            toks.insert(0, w.lower())
    return ', '.join(toks[:20])


def description(row, name, brand, cat):
    """Связное описание из атрибутов: 2-3 предложения без воды."""
    purpose = ''
    feats2 = []
    for k, v in row.items():
        kl = k.lower()
        if not v or kl in ('артикул', 'название', 'бренд', 'категория', 'цена'):
            continue
        if kl in ('назначение', 'применение', 'для кого'):
            purpose = v
        else:
            feats2.append(f'{k.lower()} {v}')
    sent1 = f'{name}' + (f' — бренд {brand}.' if brand else '.')
    sent2 = ''
    if feats2:
        sent2 = 'Выполнен: ' + ', '.join(feats2[:4]) + '.'
    purpose = re.sub(r'^(для|длЯ)\s+', '', purpose)
    sent3 = f'Подходит для: {purpose}.' if purpose else ''
    return ' '.join(s for s in (sent1, sent2, sent3) if s).strip()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('spec')
    ap.add_argument('--out', default='cards_out.xlsx')
    args = ap.parse_args()

    hdr, rows = read_rows(args.spec)
    c_name = find_col(hdr, ('название', 'наименование', 'name', 'товар'))
    c_brand = find_col(hdr, ('бренд', 'brand', 'производитель'))
    c_cat = find_col(hdr, ('категория', 'category', 'категория товара'))
    c_art = find_col(hdr, ('артикул', 'арт', 'sku', 'id'))
    if not c_name:
        sys.exit('не найдена колонка названия — переименуйте в «название»')

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'cards'
    ws.append(['артикул', 'seo_заголовок', 'описание', 'ключевые_слова',
               'характеристики_json'])
    for r in rows:
        name, brand, cat = r.get(c_name, ''), r.get(c_brand, ''), r.get(c_cat, '')
        chars = {k: v for k, v in r.items()
                 if v and k not in (c_name, c_brand, c_cat, c_art)}
        ws.append([r.get(c_art, ''), seo_title(name, brand, cat),
                   description(r, name, brand, cat), keywords(r, cat, brand),
                   json.dumps(chars, ensure_ascii=False)])
    wb.save(args.out)
    print(f'{len(rows)} карточек -> {args.out}')


if __name__ == '__main__':
    main()
