"""GOST DOC — calm, minimal Streamlit document workspace."""
from __future__ import annotations

import difflib
import hashlib
import html
import json
from datetime import date

import streamlit as st

from gostdoc.formatter import MAX_BYTES, MISIS, Profile, STANDARD, format_docx
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
from gostdoc.criteria import CATALOG, criteria_rows
from gostdoc.preview import render_docx, reference_preview, pdf_preview

st.set_page_config(
    page_title='GOST DOC · Оформление документов',
    page_icon='📄',
    layout='wide',
    initial_sidebar_state='auto',
)

st.markdown(
    """
<style>
    :root {
      --ink: #0a1e64;
      --muted: #525c69;
      --border: #dee4f0;
      --blue: #0541f0;
      --canvas: #ffffff;
    }
    html, body, [data-testid="stAppViewContainer"] {
      font-family: Helvetica, "Helvetica Neue", Arial, sans-serif; font-weight:700;
    }
    [data-testid="stAppViewContainer"] :is(h1,h2,h3,h4,h5,h6,p,label,input,textarea,button,select),
    .gd-logo, .gd-step, .gd-section, .gd-note, .gd-footer {
      font-family: Helvetica, "Helvetica Neue", Arial, sans-serif !important;
      font-weight:700 !important;
    }
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
       background:#0541f0; color:white; border:0; border-radius:8px; min-height:46px; font-weight:650;
    }
    .stButton > button[kind="primary"] p, .stDownloadButton > button[kind="primary"] p {color:#fff;}
    .stButton > button, .stDownloadButton > button {border-radius:8px; min-height:43px; font-weight:600;}
    button[data-baseweb="tab"] {font-weight:600; padding-left:18px; padding-right:18px;}
    .gd-logo {display:flex; align-items:center; gap:9px; font-weight:800; letter-spacing:-.045em; color:#0a1e64; font-size:1.35rem; padding-bottom: 15px;}
    .gd-logo-icon {display:grid; place-items:center; width:34px; height:34px; background:#f1f2fe; border-radius:10px; color:#0541f0;}
    .gd-hero {position:relative; overflow:hidden; background:#0a1e64; border:1px solid #0a1e64; border-radius:12px; padding:28px 32px; color:#fff; margin: 0 0 26px;}
    .gd-eyebrow {display:inline-flex; align-items:center; gap:7px; color:#d1defe; font-weight:700; letter-spacing:.12em; font-size:.68rem; text-transform:uppercase; padding-bottom:15px;}
    .gd-hero h1 {color:#fff!important; font-weight:760; letter-spacing:-.055em; font-size:clamp(2rem,4vw,3.2rem); line-height:1.13; margin:0 0 13px;}
    .gd-hero p {color:#c8d5ee; max-width:640px; margin:0; font-size:1rem; line-height:1.7;}
    .gd-badge {display:inline-flex; border-radius:999px; padding:5px 10px; background:#eef4ff; border:1px solid #d9e6ff; color:#0541f0; font-size:.72rem; font-weight:750; letter-spacing:.015em;}
    .gd-step {background:#fff; border:1px solid #dee4f0; border-radius:10px; padding:16px 18px; min-height:87px; margin-bottom:11px;}
    .gd-step-num {font-size:.65rem; color:#0541f0; letter-spacing:.12em; font-weight:800; text-transform:uppercase;}
    .gd-step-label {color:#0a1e64; font-size:.95rem; font-weight:750; margin-top:4px;}
    .gd-section {margin-top:15px; margin-bottom:11px; font-size:.8rem; letter-spacing:.1em; text-transform:uppercase; font-weight:800; color:#77839b;}
    .gd-note {color:#667791; font-size:.84rem; line-height:1.6;}
    .gd-footer {border-top:1px solid #dee4f0; padding-top:17px; margin-top:30px; color:#718099; font-size:.75rem;}
    @media (max-width: 700px) {
      .block-container {padding-top:1rem;}
      .gd-hero {padding:25px 22px; border-radius:12px;}
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
    '<div class="gd-hero"><div class="gd-eyebrow">GOST DOC / СТУДЕНЧЕСКИЕ ОТЧЁТЫ</div>'
    '<h1>Оформление отчёта<br>по методичке МИСИС.</h1>'
    '<p>Проверьте свой DOCX или исправьте технические параметры оформления. '
    'Финальная вёрстка — под вашим контролем.</p></div>',
    unsafe_allow_html=True,
)

with st.sidebar:
    st.markdown('<div class="gd-logo"><span class="gd-logo-icon">▤</span> GOST DOC</div>', unsafe_allow_html=True)
    st.caption('НЕЗАВИСИМЫЙ СТУДЕНЧЕСКИЙ СЕРВИС')
    mode = st.radio('Профиль оформления', ['МИСИС · методичка 2024', 'ГОСТ 7.32–2017', 'Свой профиль'], index=0)
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
        default = imported or MISIS
        key = hashlib.sha256(imported_file.getvalue()).hexdigest()[:8] if imported else 'manual'
        font = st.selectbox('Шрифт', ['Times New Roman', 'Arial', 'Calibri'],
                            index=['Times New Roman', 'Arial', 'Calibri'].index(default.font), key=f'font_{key}')
        size = st.number_input('Размер текста, пт', 10.0, 18.0, float(default.size), step=.5, key=f'size_{key}')
        spacing = st.number_input('Межстрочный интервал текста', 1.0, 2.0, float(default.line_spacing), step=.05, key=f'spacing_{key}')
        indent = st.number_input('Красная строка, см', 0.0, 2.5, float(default.first_line_cm), step=.05, key=f'indent_{key}')
        with st.expander('Параметры таблиц'):
            table_size = st.number_input('Размер текста таблиц, пт', 8.0, 18.0,
                                         float(default.for_table().size), step=.5, key=f'table_size_{key}')
            table_spacing = st.number_input('Межстрочный интервал таблиц', 1.0, 2.0,
                                            float(default.for_table().line_spacing), step=.05, key=f'table_spacing_{key}')
        with st.expander('Поля страницы (см)'):
            l = st.number_input('Левое', .5, 5., float(default.left_cm), .1, key=f'l_{key}')
            r = st.number_input('Правое', .5, 5., float(default.right_cm), .1, key=f'r_{key}')
            t = st.number_input('Верхнее', .5, 5., float(default.top_cm), .1, key=f't_{key}')
            b = st.number_input('Нижнее', .5, 5., float(default.bottom_cm), .1, key=f'b_{key}')
        cfg = Profile(font, float(size), float(spacing), float(indent), float(l), float(r), float(t), float(b),
                      ruleset='custom', table_size=float(table_size), table_line_spacing=float(table_spacing))
        st.caption('Свой профиль — ваши параметры, а не фиксированные требования МИСИС.')
    else:
        cfg = MISIS if mode.startswith('МИСИС') else STANDARD
        st.markdown('<span class="gd-badge">Профиль оформления активен</span>', unsafe_allow_html=True)
        st.markdown('**Times New Roman · 12 пт**')
        st.caption('A4 · интервал 1,5 · красная строка 1,25 см')
        st.caption('Поля: 30 / 15 / 20 / 20 мм')
        st.caption(f'Таблицы: {cfg.for_table().size:g} пт · интервал {cfg.for_table().line_spacing:g} · слева, без красной строки')
        if cfg.ruleset == 'misis':
            st.caption('Подписи таблиц: 12 пт / 1,5. Перед текстом после таблицы: 6 пт.')
    st.divider()
    st.download_button('↓ Скачать профиль JSON', profile_to_json(cfg), 'gost_doc_profile.json',
                       'application/json', use_container_width=True)
    st.caption('Базовый профиль фиксирует выбранные параметры оформления, но не заменяет полного нормоконтроля.')


steps = st.columns(3)
for col, n, text in zip(steps, ('01 · ЗАГРУЗКА', '02 · ПРОВЕРКА', '03 · РЕЗУЛЬТАТ'),
                        ('Загрузи документ', 'Проверь оформление', 'Скачай результат')):
    with col:
        st.markdown(
            f'<div class="gd-step"><div class="gd-step-num">{n}</div>'
            f'<div class="gd-step-label">{text}</div></div>',
            unsafe_allow_html=True,
        )

workspace, review, preferences = st.tabs(['◉  Рабочее пространство', '▤  Аудит оформления', '⚙  Профиль и помощь'])

with workspace:
    operation = st.radio('Режим работы', ['Исправить оформление', 'Только проверить'], horizontal=True)
    check_only = operation == 'Только проверить'
    st.markdown('<div class="gd-section">01 / Исходные файлы</div>', unsafe_allow_html=True)
    uploads = st.file_uploader(
        'Перетащи сюда файлы .docx или выбери их с компьютера',
        type=['docx'],
        accept_multiple_files=True,
        help=f'До {MAX_BATCH_FILES} файлов за раз, каждый не больше 15 МБ.',
    )
    st.caption('Документы обрабатываются на сервере приложения. Не загружайте конфиденциальные файлы без согласования.')
    if check_only:
        st.info('Проверка самостоятельного оформления: исходный файл не изменяется. Замечания можно скачать в CSV.')
    valid_files = []
    output_names = []
    if uploads:
        try:
            valid_files = read_uploads(uploads)
        except (ValueError, OSError, KeyError) as error:
            st.error(str(error))
        if valid_files:
            if not check_only:
                with st.expander('Имена готовых файлов', expanded=True):
                    st.caption('Можно указать своё имя. По методичке: фамилия_группа_дисциплина_работа№N.')
                    for index, (name, data) in enumerate(valid_files):
                        requested = st.text_input(f'Имя итогового файла · {name}',
                                                 value=safe_docx_filename(name, custom_name=name if cfg.ruleset == 'misis' else None),
                                                 key=f'output_name_{index}_{hashlib.sha256(data).hexdigest()[:12]}')
                        output_names.append(safe_docx_filename(name, custom_name=requested))
                        st.caption('Будет скачан: ' + output_names[-1])
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

    st.markdown('<div class="gd-section">02 / Проверка и сравнение</div>', unsafe_allow_html=True)
    reference = st.file_uploader('Эталон титульной страницы · DOCX или PDF', type=['docx', 'pdf'],
                                 key='cover_reference',
                                 help='Покажем первую страницу эталона рядом с исходной. ФИО, тема и содержание эталона не копируются в вашу работу.')
    reference_bytes = reference.getvalue() if reference is not None else None
    compare_key = hashlib.sha256((valid_files[0][1] if valid_files else b'') + (reference_bytes or b'')).hexdigest()
    if st.button('Сравнить титульные страницы', disabled=not valid_files or reference is None):
        try:
            with st.spinner('Строим настоящие страницы документов…'):
                original_view = render_docx(valid_files[0][1])
                reference_view = reference_preview(reference_bytes, is_pdf=reference.name.lower().endswith('.pdf'))
            st.session_state['gd_comparison'] = {'key': compare_key, 'original': original_view, 'reference': reference_view}
        except (ValueError, OSError) as error:
            st.error(str(error))
    comparison = st.session_state.get('gd_comparison')
    if comparison and comparison['key'] == compare_key:
        left, right = st.columns(2)
        with left:
            st.markdown('**Исходный документ · первая страница**')
            st.image(comparison['original']['images'][0], use_container_width=True)
        with right:
            st.markdown('**Загруженный эталон · первая страница**')
            st.image(comparison['reference']['images'][0], use_container_width=True)
        st.caption('Это рендер LibreOffice/PDF. Разметка может отличаться от Microsoft Word при замене шрифтов на сервере.')
    if not check_only and st.button('Проверить исходное оформление', disabled=not valid_files):
        source_issues = []
        for name, data in valid_files:
            source_issues.extend({**item, 'Место': f"{name} · {item['Место']}"} for item in cached_audit(data, cfg))
        st.session_state['gd_source_audit'] = {'key': hashlib.sha256(json.dumps(
            {'files': [hashlib.sha256(data).hexdigest() for _, data in valid_files], 'profile': cfg.summary()}, sort_keys=True).encode()).hexdigest(),
            'issues': source_issues}
    source_key = hashlib.sha256(json.dumps({'files': [hashlib.sha256(data).hexdigest() for _, data in valid_files],
                                           'profile': cfg.summary()}, sort_keys=True).encode()).hexdigest()
    source_audit = st.session_state.get('gd_source_audit')
    if source_audit and source_audit['key'] == source_key and not check_only:
        explain_issues(source_audit['issues'])
    st.markdown('**Параметры исправления**')
    col_a, col_b = st.columns(2)
    batch = len(valid_files) > 1
    with col_a:
        title = st.checkbox('Добавить учебный титульный лист', value=False, disabled=batch or check_only,
                            help='В пакетном режиме титульный лист выключен: каждому файлу нужны свои данные.')
        toc = st.checkbox('Создать содержание Word', value=False, disabled=check_only)
    with col_b:
        detect = st.checkbox('Распознавать заголовки автоматически', value=False, disabled=check_only,
                             help='Может принять обычный текст за заголовок. Безопаснее использовать стили Word.')
        tables = st.checkbox('Оформлять текст в таблицах', value=True, disabled=check_only,
                             help='В профиле МИСИС также исправляются ширина, заливка, соседние подписи и интервал после таблицы.')
    title = title and not batch and not check_only

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
        'detect': detect, 'tables': tables, 'cover': cover, 'operation': operation,
    }, ensure_ascii=False, sort_keys=True).encode()).hexdigest()

    if check_only and st.button('Проверить оформление', type='primary', use_container_width=True,
                                 disabled=not valid_files):
        checked = []
        with st.spinner('Проверяем параметры исходных документов…'):
            for name, data in valid_files:
                checked.extend({**issue, 'Место': f"{name} · {issue['Место']}"}
                               for issue in cached_audit(data, cfg))
        st.session_state['gd_check'] = {'signature': signature, 'issues': checked}
    check = st.session_state.get('gd_check')
    if check_only and check and check['signature'] == signature:
        st.markdown('### Результат проверки')
        counts = audit_counts(check['issues'])
        a, b, c = st.columns(3)
        a.metric('Ошибки', counts['errors'])
        b.metric('Замечания', counts['notes'])
        c.metric('Проверить вручную', counts['manual'])
        explain_issues(check['issues'])
        st.download_button('↓ Скачать проверку CSV', audit_csv(check['issues']),
                           'GOST_DOC_check.csv', 'text/csv', use_container_width=True)
        st.caption('Исходный документ не изменён. Проверка техническая и не заменяет просмотр в Word.')

    if not check_only and st.button('Оформить документы' if batch else 'Оформить документ',
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
                                   output_names[0],
                                   'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                                   type='primary', use_container_width=True)
            else:
                package = zip_results([(output_names[i], item['output']) for i, item in enumerate(results)],
                                      report=report, custom_names=True)
                st.download_button('↓ Скачать все документы ZIP', package, 'GOST_DOC_results.zip',
                                   'application/zip', type='primary', use_container_width=True)
                with st.expander('Скачать документы по отдельности'):
                    for i, item in enumerate(results):
                        st.download_button(f"↓ {output_names[i]}", item['output'],
                                           output_names[i],
                                           'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                                           key=f'file_download_{i}')
            st.download_button('↓ Скачать отчёт проверки CSV', report, 'GOST_DOC_audit.csv',
                               'text/csv', use_container_width=True)
            warnings = list(dict.fromkeys(w for r in results for w in r['warnings']))
            if warnings:
                with st.expander(f'Что проверить вручную · {len(warnings)} пункт(ов)'):
                    for warning in warnings:
                        st.write('• ' + warning)
            first = results[0]
            preview_key = hashlib.sha256(first['output']).hexdigest()
            if st.button('Показать страницы исправленного документа'):
                try:
                    with st.spinner('Рендерим исправленный DOCX…'):
                        rendered = render_docx(first['output'])
                    st.session_state['gd_result_preview'] = {'key': preview_key, **rendered}
                except (ValueError, OSError) as error:
                    st.error(str(error))
            rendered = st.session_state.get('gd_result_preview')
            if rendered and rendered['key'] == preview_key:
                page_number = st.number_input('Страница исправленного документа', 1, rendered['page_count'], 1)
                image = (rendered['images'][0] if page_number == 1 else
                         pdf_preview(rendered['pdf'], start_page=page_number - 1)['images'][0])
                st.image(image, caption=f"Исправленный документ · страница {page_number} из {rendered['page_count']}", use_container_width=True)
                st.caption('Настоящий рендер LibreOffice. Проверьте финальную вёрстку в Word перед сдачей.')
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
    elif job and not check_only:
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
    st.markdown('### Критерии из методички')
    st.caption(f"{CATALOG['source']['file']} · {CATALOG['source']['year']} · {CATALOG['source']['pages']} страниц")
    st.write(CATALOG['source']['grading'])
    st.dataframe(criteria_rows(), hide_index=True, use_container_width=True)
    st.caption('Каждое требование связано со страницей PDF. Неавтоматизированный критерий остаётся непроверенным.')
    st.markdown('### Профиль оформления')
    st.write('Здесь можно сохранить настройки для повторного использования и посмотреть действующие параметры.')
    c1, c2 = st.columns([2, 1])
    with c1:
        st.json(cfg.summary())
        st.caption('МИСИС: основной текст 12 пт / 1,5; таблицы 10 пт / 1. '
                   'ГОСТ и свой профиль — отдельные наборы параметров.')
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
    st.write('**Аудит.** Проверяются параметры текста и таблиц, подписи непосредственно перед таблицами и интервал после них. '
             'Смысл шапки, продолжения таблиц, контекстные интервалы заголовков, автоматическая нумерация, '
             'титул кафедры, формулы, источники и реальная вёрстка требуют ручной проверки. Полное соответствие не гарантируется.')
    st.write('**Безопасность.** В публичном режиме DOCX загружается на сервер приложения. Не используй его для конфиденциальных документов без соответствующих мер защиты.')
    st.markdown('[Открыть текст ГОСТ 7.32–2017](https://ifap.ru/library/gost/7322017.pdf)')
    st.caption('Исходный код и командная разработка: GitHub · Python · Streamlit.')

st.markdown(
    '<div class="gd-footer">GOST DOC · Форматирование без обещания автоматического полного нормоконтроля. '
    'Проверяй готовый документ в Word перед сдачей.</div>',
    unsafe_allow_html=True,
)
