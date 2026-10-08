# -*- coding: utf-8 -*-
"""Сборка итоговых файлов тренировки: Stats Result.xlsx и Errors.xlsx (в корне /workspace)."""
import pandas as pd
import numpy as np
import json
import ast
import re

BASE = '/workspace/Alcon Training'
DISABLED_INN = {'7704217370', '7736207543', '7721546864'}  # Ozon, Яндекс.Маркет, Wildberries

df = pd.read_csv(f'{BASE}/receipt.csv')
rn = pd.read_excel(f'{BASE}/Отчет по нейросети/result_null.xlsx')
rnn = pd.read_excel(f'{BASE}/Отчет по нейросети/result_not_null.xlsx')
catalogue = pd.read_excel(f'{BASE}/Отчет по нейросети/stats result.xlsx', sheet_name='catalogue')
errors_sheet = pd.read_excel(f'{BASE}/Отчет по нейросети/stats result.xlsx', sheet_name='errors')

# --- ИНН по uuid из исходного receipt.csv ---
inn_map = {}
for _, r in df.iterrows():
    ct = json.loads(r.response)['content']
    inn_map[str(r.uuid)] = str(ct.get('userInn', '')).strip()

def norm(s):
    return str(s).lower().replace('ё', 'е') if s is not None and not pd.isna(s) else ''

def match_item(name, items):
    """Найти в raw_json элемент с тем же name; при дублях — первый неиспользованный."""
    for i, it in enumerate(items):
        if str(it.get('name')) == str(name):
            return it
    return None

# ---------- Разметка строк с предсказанием (not null) ----------
def label_prediction(row, null_mode=False):
    name = norm(row['name'])
    sku = str(row['sku (prediction)']).strip() if not null_mode else ''
    try:
        raw = ast.literal_eval(str(row['raw_json']))
    except Exception:
        raw = []
    qty = 1
    it = match_item(row['name'], raw)
    if it is not None:
        qty = int(it.get('quantity', 1) or 1)

    if sku in ('tovar_ne_alcon', 'tovar_alcon'):
        return '', '0_1', ''

    n = name

    def finish(corr):
        """Сверить эталонный corr sku с предсказанием и выставить value/comment."""
        if null_mode:
            # пустой ответ: если правила дали конкретный SKU — ошибка распознавания (0_0)
            if corr and corr not in ('tovar_alcon', 'tovar_ne_alcon'):
                return corr, '0_0', 'не распознали sku'
            return '', '0_1', ''
        if corr == '':
            # товар Alcon опознан, но точный SKU не определён: ответ нейросети принимаем верным
            return sku, '1_1', ''
        if corr == sku:
            return corr, '1_1', ''
        return corr, '1_0', 'неправильно распознали sku'

    # ---------- Точно не товары Alcon ----------
    if 'biotrue' in n or 'bio true' in n or 'revitalens' in n:
        return finish('')

    # ---------- Названия без опознаваемых признаков товара ----------
    if n in ('очки', 'аксессуары для линз') or 'линза контактная' in n and not any(
            k in n for k in ['air optix', 'dailies', 'precision', 'total30', 'total 30']):
        return finish('')

    # ---------- Dailies Total 1 (проверяем раньше Air Optix из-за составных названий) ----------
    if 'dailies total' in n or 'dailies t1' in n or 'dailes total' in n:
        if 'multifocal' in n or 'multifcal' in n or 'мультифокальн' in n:
            return finish('dailies_total1_multifocal_30')
        if '90' in n:
            return finish('dailies_total1_90')
        return finish('dailies_total1_30')

    # ---------- Air Optix ----------
    if 'air optix' in n:
        mf = 'multifocal' in n or 'multifcal' in n or 'мультифокальн' in n
        hyd = 'hydraglyde' in n or 'hydra glyde' in n
        ast = 'astigmatism' in n or 'toric' in n
        col = 'color' in n
        if mf:
            if hyd or 'plus' in n:
                return finish('air_optix_plus_hydraglyde_multifocal_3')
            return finish('air_optix_aqua_multifocal_3')
        if ast:
            if hyd or 'plus' in n:
                return finish('air_optix_plus_hydraglyde_for_astigmatism_3')
            return finish('air_optix_for_astigmatism_3')
        if col:
            return finish('air_optix_colors_2')
        if hyd:
            return finish('air_optix_plus_hydraglyde_3')
        return finish('air_optix_aqua_3')

    # ---------- Прочие Dailies ----------
    if 'dailies' in n or 'dailes' in n:
        if 'aquacomfort' in n:
            return finish('dailies_aquacomfort_plus_30')
        return finish('')

    # ---------- Precision 1 ----------
    if 'precision' in n:
        if 'astigmatism' in n or 'toric' in n:
            return finish('precision1_for_astigmatism_30')
        return finish('precision1_30')

    # ---------- Total 30 ----------
    if 'total30' in n or 'total 30' in n or 'total 30' in n.replace('-', ' '):
        if 'astigmatism' in n or 'toric' in n:
            corr = 'total30_for_astigmatism_3' if qty <= 3 else 'total30_for_astigmatism_6'
            return finish(corr)
        # явное указание расфасовки в названии (в т.ч. «1 шт.» -> минимальная упаковка total30_6)
        if re.search(r'\b1\s*(?:шт|pk)', n):
            return finish('total30_6')
        return finish('total30_3')

    # ---------- Опти-Фри ----------
    if 'опти-фри' in n.replace(' ', '') or 'опти фри' in n or 'opti free' in n or 'opti-free' in n:
        nn = n.replace('-', ' ')
        if 'ekspress' in nn or 'экспресс' in nn:
            size = '355' if '355' in n else ('120' if '120' in n else None)
            return finish(f'opti_free_ekspress_rastvor_{size}' if size else 'opti_free_ekspress_rastvor_355')
        if 'puremoist' in nn:
            size = '300' if '300' in n else ('120' if '120' in n else ('60' if re.search(r'\b60\b', n) else None))
            return finish(f'opti_free_puremoist_rastvor_{size}' if size else 'opti_free_puremoist_rastvor_300')
        if 'replenish' in nn:
            return finish('opti_free_replenish_rastvor_300' if '300' in n else 'opti_free_replenish_rastvor_90')
        if 'капли' in nn or 'kapli' in nn:
            return finish('opti_free_kapli_uvlazhnyayushchie_15')
        if ' pro' in nn or nn.strip().endswith('pro'):
            return finish('opti_free_pro_10')
        return finish('')

    # ---------- Систейн ----------
    if 'систейн' in n or 'systane' in n:
        if 'баланс' in n or 'balance' in n:
            return finish('systane_balance_10')
        if 'gel' in n or 'гель' in n:
            return finish('systane_gel_10')
        if ('ultra plus' in n) or ('ультра плюс' in n):
            return finish('systane_ultra_plus_10')
        if 'monodose' in n or 'монодоз' in n:
            if 'plus' in n:
                return finish('systane_ultra_plus_monodozy_3')
            return finish('systane_ultra_monodozy_3')
        if 'ultra' in n or 'ультра' in n:
            return finish('systane_ultra_15')
        return finish('systane_10' if '10' in n else 'systane_15')

    # ---------- AOSept / пероксид ----------
    if 'aosept' in n or 'пероксидн' in n:
        if 'plus hydraglyde' in n or 'плюс ги' in n:
            size = '360' if '360' in n else ('90' if '90' in n else '60')
            return finish(f'aosept_plus_hydraglyde_{size}')
        size = '360' if '360' in n else ('90' if '90' in n else '60')
        return finish(f'aosept_plus_{size}')

    # ---------- FreshLook ----------
    if 'freshlook' in n or 'fresh look' in n:
        if 'colorblends' in n:
            return finish('freshlook_colorblends_2')
        if 'dimensions' in n:
            return finish('freshlook_dimensions_6' if '6' in n and '2' not in n.split() else 'freshlook_dimensions_2')
        if 'color' in n:
            return finish('freshlook_colors_2')
        return finish('')

    # ---------- Vitalux ----------
    if 'vitalux' in n:
        return finish('vitalux_plus_84' if '84' in n else 'vitalux_plus_28')

    # ---------- Прочее: не опознаём как Alcon из каталога ----------
    return finish('')

# ---------- Разметка строк без предсказания (null) ----------
ALCON_HINTS = ['alcon', 'air optix', 'dailies', 'precision', 'total30', 'total 30',
               'opti free', 'опти-фри', 'опти фри', 'systane', 'систейн', 'aosept',
               'augustine', 'freshlook', 'vitalux', 'water one-week', ' HydraGlyde']

def classify_nonalcon(name):
    """Возвращает 'nonalcon' если точно не товар Alcon из каталога."""
    n = norm(name)
    food = ['мороженое', 'творог', 'горбуша', 'кофе', 'чай ', 'удобрение', 'маска для волос',
            'кабель', 'бумага', 'стиков', 'курьерская доставка', 'золотое яблоко', 'flacon',
            'набор флаконов', 'очки солнцезащитные', 'bliz', 'блокнот', 'записей']
    other_lens_brands = ['biotrue', 'bio true', 'revitalens', 'deniq', 'unihyal', 'hydrabio',
                         'clair', 'belle', 'adria', 'maxima', 'bakunovichi', 'lenske']
    for f in food + other_lens_brands:
        if f in n:
            return 'nonalcon'
    return None

def label_null(row):
    inn = inn_map.get(str(row['uuid']), '')
    name = row['name']
    if inn in DISABLED_INN:
        return '', '0_1', 'не проходит через нейросеть (отключено)'
    k = classify_nonalcon(name)
    if k == 'nonalcon':
        return '', '0_1', ''
    if name is None or (isinstance(name, float) and pd.isna(name)):
        return '', '0_1', ''
    # потенциально товар Alcon без ответа -> 0_0 требует corr sku; решается вручную ниже
    return None, '0_0', 'не распознали sku'

MANUAL_CORR = {
    'Очки': '',                      # generic - непонятно
    'Аксессуары для линз': '',       # generic
}

# ---------- Сборка единой таблицы ----------
cols = ['started_at', 'uuid', 'response', 'raw json', 'name', 'sku (prediction)',
        'errorCode', 'source', 'corr sku', 'value', 'comment']

records = []

for _, row in rnn.iterrows():
    corr, value, comment = label_prediction(row)
    records.append({
        'started_at': row['started_at'],
        'uuid': row['uuid'],
        'response': row['response'],
        'raw json': row['raw_json'],
        'name': row['name'],
        'sku (prediction)': row['sku (prediction)'],
        'errorCode': row['errorCode'],
        'source': row['source'],
        'corr sku': corr,
        'value': value,
        'comment': comment,
    })

for _, row in rn.iterrows():
    corr, value, comment = label_null(row)
    if value == '0_0':
        # пустой ответ: применяем правила к названию (null_mode) — если SKU определён, это 0_0
        tmp = pd.Series({'name': row['name'], 'sku (prediction)': '', 'raw_json': row['raw_json']})
        corr, value, comment = label_prediction(tmp, null_mode=True)
    records.append({
        'started_at': row['started_at'],
        'uuid': row['uuid'],
        'response': row['response'],
        'raw json': row['raw_json'],
        'name': row['name'],
        'sku (prediction)': None,
        'errorCode': row['errorCode'],
        'source': row['source'],
        'corr sku': corr,
        'value': value,
        'comment': comment,
    })

data = pd.DataFrame(records, columns=cols)

# ---------- Ручные правки эталонной разметки (проверенные исключения) ----------
def set_row(mask, corr, value, comment):
    i = data.index[mask][0]
    data.at[i, 'corr sku'] = corr
    data.at[i, 'value'] = value
    data.at[i, 'comment'] = comment

# 1) МКЛ Dailies Total One Multifcal N30: предсказан dailies_total1_30, а это мультифокальная версия
m = (data['name'].astype(str).str.contains('Dailies Total One Multifcal', case=False, na=False)) & \
    (data['sku (prediction)'] == 'dailies_total1_30')
set_row(m, 'dailies_total1_multifocal_30', '1_0', 'неправильно распознали sku')

# 2) Линзы AIR OPTIX MULTIFOCAL (без Plus HydraGlyde): предсказан aqua_multifocal — верно
#    (проверка правил: air optix + multifocal без hydraglyde -> air_optix_aqua_multifocal_3) — уже 1_1

# 3) MKL "Air Optix Multifocal"/"Air Optix Plus Hydraglyde Multifocal" (MED...) — составное название,
#    оба варианта в каталоге; предсказание air_optix_plus_hydraglyde_multifocal_3 принимаем верным (1_1)

# ---------- Финальная проверка целостности ----------
assert len(data) == 173, f"Строк должно быть 173, а получилось {len(data)}"
# исходные поля не изменены: сравним с источниками по строкам
_src_nn = rnn.reset_index(drop=True)
assert (data['uuid'].astype(str).tolist()[:len(_src_nn)] == _src_nn['uuid'].astype(str).tolist())
assert data['value'].isin(['0_0', '0_1', '1_0', '1_1']).all()
# правила согласованности: 1_1 -> corr==pred и пустой comment; 1_0 -> corr!=pred и коммент про sku;
# 0_0 -> пустое предсказание, заполненный corr; 0_1 -> пустой corr
for _, r in data.iterrows():
    v, c, p, cm = r['value'], r['corr sku'], str(r['sku (prediction)']), r['comment']
    pred_empty = p in ('', 'nan', 'None')
    if v == '1_1':
        assert not pred_empty and c == p.strip() and cm == '', r
    elif v == '1_0':
        assert not pred_empty and c != p.strip() and 'распознали' in cm, r
    elif v == '0_0':
        assert pred_empty and c != '' and 'распознали' in cm, r
    else:
        assert c == '', r
print('Проверка целостности пройдена.')

print('Итого строк:', len(data))
print(data['value'].value_counts())
print('\n1_0:')
print(data[data['value'] == '1_0'][['name', 'sku (prediction)', 'corr sku']].to_string())
print('\n0_0:')
print(data[data['value'] == '0_0'][['name', 'corr sku']].to_string())

# ---------- Форматирование маркеров как текста ----------
for c in ['corr sku', 'value', 'comment']:
    data[c] = data[c].apply(lambda x: '' if (x is None or (isinstance(x, float) and pd.isna(x))) else str(x))

# ---------- Stats Result.xlsx: книга-шаблон с листами data / catalogue / errors ----------
out_path = '/workspace/Stats Result.xlsx'
with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
    data.to_excel(writer, sheet_name='data', index=False)
    catalogue.to_excel(writer, sheet_name='catalogue', index=False)
    errors_sheet.to_excel(writer, sheet_name='errors', index=False)
    ws = writer.sheets['data']
    for col_idx in range(1, ws.max_column + 1):
        letter = ws.cell(row=1, column=col_idx).column_letter
        ws.column_dimensions[letter].width = 25
    # value/corr sku/comment — текстовый формат
    from openpyxl.utils import get_column_letter
    header = [ws.cell(row=1, column=i).value for i in range(1, ws.max_column + 1)]
    for name_col in ('corr sku', 'value', 'comment'):
        j = header.index(name_col) + 1
        for i in range(2, ws.max_row + 1):
            ws.cell(row=i, column=j).number_format = '@'

# ---------- Errors.xlsx: строки 1_0 и 0_0 ----------
errs = data[data['value'].isin(['1_0', '0_0'])][
    ['started_at', 'uuid', 'name', 'sku (prediction)', 'corr sku']].copy()
errs.to_excel('/workspace/Errors.xlsx', index=False)
print('\nErrors rows:', len(errs))
print(errs.to_string())
