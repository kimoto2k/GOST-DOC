"""GOST DOC — calm, minimal Streamlit document workspace."""
from __future__ import annotations

import difflib
import hashlib
import html
import json
from datetime import date

import streamlit as st

from gostdoc.formatter import MAX_BYTES, Profile, STANDARD, format_docx
from gostdoc.workflow import (
    MAX_BATCH_FILES,
    audit_counts,
    audit_csv,
    audit_document,
    demo_docx,
    inspect_document,
    profile_from_json,
    profile_to_json,
    safe_docx_filename,
    zip_results,
)

st.set_page_config(
    page_title='GOST DOC · Оформление документов',
    page_icon='📄',
    layout='wide',
    initial_sidebar_state='expanded',
)

st.markdown(
    """
<style>
    :root {
      --ink: #17243a;
      --muted: #69788e;
      --border: #e1e8f0;
      --blue: #315dcd;
      --canvas: #f7f9fc;
    }
    html, body, [class*="css"] {font-family: Inter, -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, sans-serif;}
    [data-testid="stAppViewContainer"] {background: var(--canvas);}
    [data-testid="stSidebar"] {background: #ffffff; border-right: 1px solid var(--border);}
    [data-testid="stSidebar"] .block-container {padding-top: 1.8rem;}
    .block-container {max-width: 1190px; padding-top: 2.2rem; padding-bottom: 4rem;}
    h1, h2, h3 {letter-spacing: -.035em; color: var(--ink);}
    h2 {font-size: 1.55rem !important; font-weight: 750 !important;}
    h3 {font-size: 1.12rem !important;}
    p, label {color: var(--ink);}
    div[data-testid="stFileUploader"] {background: #fff; border-radius: 18px; border: 1px solid var(--border); padding: 12px 16px;}
    div[data-testid="stFileUploader"] section {border-radius: 14px; background: #f8faff; border: 1px dashed #bdcce8;}
    div[data-testid="stVerticalBlockBorderWrapper"] {border-color: var(--border); background: #fff; border-radius: 16px;}
    div[data-testid="stMetric"] {background: white; border: 1px solid var(--border); border-radius: 14px; padding: 12px 16px;}
    div[data-testid="stMetric"] label {font-size: .85rem; color: var(--muted);}
    div[data-testid="stMetric"] [data-testid="stMetricValue"] {font-weight: 750; color: var(--ink);}
    .stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {
       background:#315dcd; color:white; border:0; border-radius:11px; min-height:46px; font-weight:650;
    }
    .stButton > button, .stDownloadButton > button {border-radius:11px; min-height:43px; font-weight:600;}
    button[data-baseweb="tab"] {font-weight:600; padding-left:18px; padding-right:18px;}
    .gd-logo {display:flex; align-items:center; gap:9px; font-weight:800; letter-spacing:-.045em; color:#192e4b; font-size:1.35rem; padding-bottom: 15px;}
    .gd-logo-icon {display:grid; place-items:center; width:34px; height:34px; background:#eaf0ff; border-radius:10px; color:#315dcd;}
    .gd-hero {position:relative; overflow:hidden; background:#172b4d; border:1px solid #28466f; border-radius:22px; padding:35px 38px; color:#fff; margin: 0 0 26px;}
    .gd-hero::after {content:""; position:absolute; width:300px; height:300px; top:-160px; right:-30px; border:60px solid rgba(164,195,255,.07); border-radius:50%; pointer-events:none;}
    .gd-eyebrow {display:inline-flex; align-items:center; gap:7px; color:#d1defe; font-weight:700; letter-spacing:.12em; font-size:.68rem; text-transform:uppercase; padding-bottom:15px;}
    .gd-hero h1 {color:#fff!important; font-weight:760; letter-spacing:-.055em; font-size:clamp(2rem,4vw,3.2rem); line-height:1.13; margin:0 0 13px;}
    .gd-hero p {color:#c8d5ee; max-width:640px; margin:0; font-size:1rem; line-height:1.7;}
    .gd-badge {display:inline-flex; border-radius:999px; padding:5px 10px; background:#eef4ff; border:1px solid #d9e6ff; color:#315dcd; font-size:.72rem; font-weight:750; letter-spacing:.015em;}
    .gd-step {background:#fff; border:1px solid #e1e8f0; border-radius:14px; padding:16px 18px; min-height:87px; margin-bottom:11px;}
    .gd-step-num {font-size:.65rem; color:#315dcd; letter-spacing:.12em; font-weight:800; text-transform:uppercase;}
    .gd-step-label {color:#1b2a44; font-size:.95rem; font-weight:750; margin-top:4px;}
    .gd-section {margin-top:15px; margin-bottom:11px; font-size:.8rem; letter-spacing:.1em; text-transform:uppercase; font-weight:800; color:#77839b;}
    .gd-note {color:#667791; font-size:.84rem; line-height:1.6;}
    .gd-footer {border-top:1px solid #e1e8f0; padding-top:17px; margin-top:30px; color:#718099; font-size:.75rem;}
    @media (max-width: 700px) {
      .block-container {padding-top:1rem;}
      .gd-hero {padding:25px 22px; border-radius:17px;}
      .gd-hero h1 {font-size:2rem;}
    }
</style>
""",
    unsafe_allow_html=True,
)


@st.cache_data(show_spinner=False, max_entries=24)
def cached_inspect(data: bytes) -> dict:
    return inspect_document(data)


@st.cache_data(show_spinner=False, max_entries=24)
def cached_audit(data: bytes, cfg: Profile) -> list[dict]:
    return audit_document(data, cfg)


def explain_issues(issues: list[dict], limit: int = 150) -> None:
    if issues:
        st.dataframe(issues[:limit], hide_index=True, use_container_width=True)
        if len(issues) > limit:
            st.caption(f'Показаны первые {limit} замечаний.')
    else:
        st.success('По реализованным техническим проверкам замечаний не найдено.')


def read_uploads(uploaded) -> list[tuple[str, bytes]]:
    items = [(file.name, file.getvalue()) for file in uploaded]
    if len(items) > MAX_BATCH_FILES:
        raise ValueError(f'За один раз можно обработать до {MAX_BATCH_FILES} документов.')
    if any(len(data) > MAX_BYTES for _, data in items):
        raise ValueError('Файл превышает допустимые 15 МБ.')
    for _, data in items:
        cached_inspect(data)  # Includes validation for valid, bounded DOCX.
    return items


st.markdown(
    '<div class="gd-hero"><div class="gd-eyebrow">● GOST DOC · DOCUMENT WORKSPACE</div>'
    '<h1>Документы в порядке.<br>Без лишней работы.</h1>'
    '<p>Загружай Word, проверяй оформление и получай готовый DOCX. '
    'Всё в одном спокойном рабочем пространстве.</p></div>',
    unsafe_allow_html=True,
)

with st.sidebar:
    st.markdown('<div class="gd-logo"><span class="gd-logo-icon">▤</span> GOST DOC</div>', unsafe_allow_html=True)
    st.caption('WORKSPACE · ВЕРСИЯ 2')
    mode = st.radio('Профиль оформления', ['ГОСТ 7.32–2017', 'Свой профиль'], index=0)
    imported = None
    if mode == 'Свой профиль':
        imported_file = st.file_uploader(
            'Импорт профиля JSON', type=['json'], key='profile_json',
            help='Можно загрузить сохранённые настройки оформления GOST DOC.'
        )
        if imported_file is not None:
            try:
                imported = profile_from_json(imported_file.getvalue())
                st.success('Профиль загружен.')
            except ValueError as error:
                st.error(str(error))
        default = imported or STANDARD
        key = hashlib.sha256(imported_file.getvalue()).hexdigest()[:8] if imported else 'manual'
        font = st.selectbox('Шрифт', ['Times New Roman', 'Arial', 'Calibri'],
                            index=['Times New Roman', 'Arial', 'Calibri'].index(default.font), key=f'font_{key}')
        size = st.number_input('Размер текста, пт', 10, 18, int(default.size), key=f'size_{key}')
        spacing = st.select_slider('Интервал', [1.0, 1.15, 1.5, 2.0], value=default.line_spacing, key=f'spacing_{key}')
        indent = st.number_input('Красная строка, см', 0.0, 2.5, float(default.first_line_cm), step=.05, key=f'indent_{key}')
        with st.expander('Поля страницы (см)'):
            l = st.number_input('Левое', .5, 5., float(default.left_cm), .1, key=f'l_{key}')
            r = st.number_input('Правое', .5, 5., float(default.right_cm), .1, key=f'r_{key}')
            t = st.number_input('Верхнее', .5, 5., float(default.top_cm), .1, key=f't_{key}')
            b = st.number_input('Нижнее', .5, 5., float(default.bottom_cm), .1, key=f'b_{key}')
        cfg = Profile(font, int(size), float(spacing), float(indent), float(l), float(r), float(t), float(b))
    else:
        cfg = STANDARD
        st.markdown('<span class="gd-badge">Профиль оформления активен</span>', unsafe_allow_html=True)
        st.markdown('**Times New Roman · 12 пт**')
        st.caption('A4 · интервал 1,5 · красная строка 1,25 см')
        st.caption('Поля: 30 / 15 / 20 / 20 мм')
    st.divider()
    st.download_button('↓ Скачать профиль JSON', profile_to_json(cfg), 'gost_doc_profile.json',
                       'application/json', use_container_width=True)
    st.caption('Базовый профиль фиксирует выбранные параметры оформления, но не заменяет полного нормоконтроля.')


steps = st.columns(3)
for col, n, text in zip(steps, ('01 · UPLOAD', '02 · CHECK', '03 · EXPORT'),
                        ('Загрузи документ', 'Проверь оформление', 'Скачай результат')):
    with col:
        st.markdown(
            f'<div class="gd-step"><div class="gd-step-num">{n}</div>'
            f'<div class="gd-step-label">{text}</div></div>',
            unsafe_allow_html=True,
        )

workspace, review, preferences = st.tabs(['◉  Рабочее пространство', '▤  Аудит оформления', '⚙  Профиль и помощь'])

with workspace:
    st.markdown('<div class="gd-section">01 / Исходные файлы</div>', unsafe_allow_html=True)
    uploads = st.file_uploader(
        'Перетащи сюда файлы .docx или выбери их с компьютера',
        type=['docx'],
        accept_multiple_files=True,
        help=f'До {MAX_BATCH_FILES} файлов за раз, каждый не больше 15 МБ.',
    )
    valid_files = []
    if uploads:
        try:
            valid_files = read_uploads(uploads)
        except (ValueError, OSError, KeyError) as error:
            st.error(str(error))
        if valid_files:
            st.caption(f'Загружено файлов: {len(valid_files)} · Максимум {MAX_BATCH_FILES} за один пакет.')
            summary = [cached_inspect(data) for _, data in valid_files]
            m1, m2, m3 = st.columns(3)
            m1.metric('Слов', f"{sum(s['words'] for s in summary):,}".replace(',', ' '))
            m2.metric('Абзацев', sum(s['paragraphs'] for s in summary))
            m3.metric('Таблиц', sum(s['tables'] for s in summary))
            if len(valid_files) == 1:
                with st.expander('Структура и предварительный просмотр', expanded=True):
                    info = summary[0]
                    if info['headings']:
                        st.markdown('**Распознанные заголовки**')
                        for heading in info['headings'][:16]:
                            st.write('　' * (heading['level'] - 1) + '• ' + heading['text'])
                    else:
                        st.caption('Заголовки не распознаны. Лучше применять стили Word «Заголовок 1–3».')
                    from gostdoc.formatter import get_document
                    src = get_document(valid_files[0][1])
                    snippet = '\n'.join(p.text for p in src.paragraphs if p.text.strip())[:5000]
                    st.text_area('Текст документа · первые 5 000 символов', snippet,
                                 height=180, disabled=True)

    st.markdown('<div class="gd-section">02 / Что сделать с документом</div>', unsafe_allow_html=True)
    col_a, col_b = st.columns(2)
    batch = len(valid_files) > 1
    with col_a:
        title = st.checkbox('Добавить учебный титульный лист', value=False, disabled=batch,
                            help='В пакетном режиме титульный лист выключен: каждому файлу нужны свои данные.')
        toc = st.checkbox('Создать содержание Word', value=False)
    with col_b:
        detect = st.checkbox('Распознавать заголовки автоматически', value=False,
                             help='Может принять обычный текст за заголовок. Безопаснее использовать стили Word.')
        tables = st.checkbox('Оформлять текст в таблицах', value=True)
    title = title and not batch

    cover = {}
    if title:
        st.info('Учебный титульный лист — упрощённый шаблон, не официальный титульный лист отчёта НИР.')
        with st.expander('Данные титульного листа', expanded=True):
            x, y = st.columns(2)
            with x:
                cover['organization'] = st.text_input('Организация', 'Министерство науки и высшего образования РФ')
                cover['university'] = st.text_input('Университет', 'НИТУ МИСИС')
                cover['department'] = st.text_input('Институт / кафедра', '')
                cover['work_type'] = st.text_input('Тип работы', 'ОТЧЁТ')
                cover['topic'] = st.text_input('Название работы', '')
            with y:
                cover['student'] = st.text_input('ФИО студента', '')
                cover['group'] = st.text_input('Группа', '')
                cover['teacher'] = st.text_input('ФИО преподавателя', '')
                cover['city'] = st.text_input('Город', 'Москва')
                cover['year'] = st.text_input('Год', str(date.today().year))

    signature = hashlib.sha256(json.dumps({
        'files': [(name, hashlib.sha256(data).hexdigest()) for name, data in valid_files],
        'profile': cfg.summary(), 'title': title, 'toc': toc,
        'detect': detect, 'tables': tables, 'cover': cover,
    }, ensure_ascii=False, sort_keys=True).encode()).hexdigest()

    if st.button('Оформить документы' if batch else 'Оформить документ',
                 type='primary', use_container_width=True, disabled=not valid_files):
        results = []
        failures = []
        with st.spinner('Применяем настройки и проверяем результат…'):
            for name, data in valid_files:
                try:
                    output, before_text, after_text, warnings = format_docx(
                        data, cfg, add_title=title, cover=cover, create_toc=toc,
                        recognize_headings=detect, format_tables=tables,
                    )
                    results.append({
                        'name': name, 'output': output, 'before_text': before_text,
                        'after_text': after_text, 'warnings': warnings,
                        'before_issues': cached_audit(data, cfg),
                        'after_issues': cached_audit(output, cfg),
                    })
                except (ValueError, OSError, KeyError, TypeError) as error:
                    failures.append(f'{name}: {error}')
        st.session_state['gd_job'] = {'signature': signature, 'results': results, 'failures': failures}

    job = st.session_state.get('gd_job')
    if job and job['signature'] == signature:
        st.markdown('<div class="gd-section">03 / Готовые файлы</div>', unsafe_allow_html=True)
        for msg in job['failures']:
            st.error(msg)
        results = job['results']
        if results:
            st.success(f'Готово: {len(results)} документ(ов). Проверь финальную верстку в Word.')
            err_before = sum(audit_counts(r['before_issues'])['errors'] for r in results)
            err_after = sum(audit_counts(r['after_issues'])['errors'] for r in results)
            c1, c2 = st.columns(2)
            c1.metric('Ошибок до обработки', err_before)
            c2.metric('Ошибок после обработки', err_after, delta=err_after - err_before, delta_color='inverse')
            all_audit = []
            for item in results:
                for entry in item['after_issues']:
                    all_audit.append({**entry, 'Место': f"{item['name']} · {entry['Место']}"})
            report = audit_csv(all_audit)
            if len(results) == 1:
                item = results[0]
                st.download_button('↓ Скачать готовый DOCX', item['output'],
                                   safe_docx_filename(item['name']),
                                   'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                                   type='primary', use_container_width=True)
            else:
                package = zip_results([(item['name'], item['output']) for item in results], report=report)
                st.download_button('↓ Скачать все документы ZIP', package, 'GOST_DOC_results.zip',
                                   'application/zip', type='primary', use_container_width=True)
                with st.expander('Скачать документы по отдельности'):
                    for i, item in enumerate(results):
                        st.download_button(f"↓ {item['name']}", item['output'],
                                           safe_docx_filename(item['name']),
                                           'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                                           key=f'file_download_{i}')
            st.download_button('↓ Скачать отчёт проверки CSV', report, 'GOST_DOC_audit.csv',
                               'text/csv', use_container_width=True)
            warnings = list(dict.fromkeys(w for r in results for w in r['warnings']))
            if warnings:
                with st.expander(f'Что проверить вручную · {len(warnings)} пункт(ов)'):
                    for warning in warnings:
                        st.write('• ' + warning)
            with st.expander('Сравнение текста первого документа'):
                first = results[0]
                left, right = st.columns(2)
                with left:
                    st.markdown('**До**')
                    st.text_area('До', first['before_text'][:4000], height=230,
                                 disabled=True, label_visibility='collapsed')
                with right:
                    st.markdown('**После**')
                    st.text_area('После', first['after_text'][:4000], height=230,
                                 disabled=True, label_visibility='collapsed')
                st.caption('Это сравнение текста, не визуального расположения страниц.')
                diff = '\n'.join(difflib.unified_diff(first['before_text'].splitlines(),
                                                     first['after_text'].splitlines(),
                                                     fromfile='До', tofile='После', lineterm=''))
                st.code((diff or 'Текст не изменён.')[:9000], language='diff')
    elif job:
        st.caption('Файл или настройки изменились. Для получения нового результата нажми «Оформить документ».')
    if not uploads:
        st.caption('Можно попробовать приложение без своих файлов: скачай демонстрационный DOCX на вкладке «Профиль и помощь».')

with review:
    st.markdown('### Аудит оформления')
    st.write('Проверь поля, размер и тип шрифта, интервалы и другие доступные технические параметры до обработки.')
    if not uploads:
        st.info('Сначала загрузи DOCX на вкладке «Рабочее пространство».')
    else:
        try:
            selected = read_uploads(uploads)
        except (ValueError, OSError, KeyError) as error:
            st.error(str(error))
            selected = []
        if selected:
            names = [name for name, _ in selected]
            chosen = st.selectbox('Документ для проверки', range(len(names)), format_func=lambda index: names[index])
            issues = cached_audit(selected[chosen][1], cfg)
            counts = audit_counts(issues)
            a, b, c = st.columns(3)
            a.metric('Ошибки', counts['errors'])
            b.metric('Замечания', counts['notes'])
            c.metric('Проверить вручную', counts['manual'])
            categories = sorted({r['Категория'] for r in issues})
            category = st.selectbox('Категория', ['Все'] + categories)
            filtered = issues if category == 'Все' else [r for r in issues if r['Категория'] == category]
            explain_issues(filtered)
            st.download_button('↓ Скачать аудит CSV', audit_csv(issues),
                               'GOST_DOC_audit.csv', 'text/csv')
            st.caption('Отсутствие замечаний не означает полного соответствия ГОСТ. Сложные элементы нужно проверять в Word.')

with preferences:
    st.markdown('### Профиль оформления')
    st.write('Здесь можно сохранить настройки для повторного использования и посмотреть действующие параметры.')
    c1, c2 = st.columns([2, 1])
    with c1:
        st.json(cfg.summary())
        st.caption('Стандартный профиль сервиса: Times New Roman 12 пт; поля 30/15/20/20 мм; полуторный интервал.')
    with c2:
        st.download_button('↓ Сохранить профиль JSON', profile_to_json(cfg),
                           'GOST_DOC_profile.json', 'application/json', use_container_width=True)
        st.download_button('↓ Тестовый документ DOCX', demo_docx(),
                           'GOST_DOC_demo.docx',
                           'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                           use_container_width=True)
    st.divider()
    st.markdown('### Что важно знать')
    st.write('**Содержание.** Используй стили Word «Заголовок 1–3». После открытия результата обнови поле содержания клавишей F9 или через контекстное меню.')
    st.write('**Титульный лист.** Это учебный шаблон. Для официального отчёта НИР требуется соответствующая форма по стандарту и требованиям организации.')
    st.write('**Аудит.** Проверка не анализирует все требования ГОСТ, смысл текста, библиографию, подписи и формулы, либо верстку страницы как Microsoft Word.')
    st.write('**Безопасность.** В публичном режиме DOCX загружается на сервер приложения. Не используй его для конфиденциальных документов без соответствующих мер защиты.')
    st.markdown('[Открыть текст ГОСТ 7.32–2017](https://ifap.ru/library/gost/7322017.pdf)')
    st.caption('Исходный код и командная разработка: GitHub · Python · Streamlit.')

st.markdown(
    '<div class="gd-footer">GOST DOC · Форматирование без обещания автоматического полного нормоконтроля. '
    'Проверяй готовый документ в Word перед сдачей.</div>',
    unsafe_allow_html=True,
)
