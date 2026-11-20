"""GOST DOC: optional public Streamlit UI and local usage."""
import difflib
import hashlib
import json
from datetime import date

import streamlit as st

from gostdoc.formatter import Profile, STANDARD, format_docx, diagnostics, get_document, MAX_BYTES

st.set_page_config(page_title='GOST DOC', page_icon='📄', layout='wide')
st.markdown('''
<style>
.block-container{max-width:1050px;padding-top:1.8rem}
h1{letter-spacing:-.035em}
[data-testid="stMetric"]{border:1px solid #e4e9f0;padding:14px;border-radius:12px}
</style>
''', unsafe_allow_html=True)
st.title('📄 GOST DOC')
st.caption('Оформление Word-документов по профилю ГОСТ 7.32–2017 · версия для совместной разработки')

with st.sidebar:
    st.subheader('Стандарт оформления')
    mode = st.radio('Режим', ['ГОСТ 7.32–2017 · TNR 12', 'Свои параметры'], label_visibility='collapsed')
    if mode.startswith('ГОСТ'):
        cfg = STANDARD
        st.write('**Основной текст:** Times New Roman, 12 пт')
        st.write('**Интервал:** 1,5 · красная строка 1,25 см')
        st.write('**Поля:** 30 / 15 / 20 / 20 мм')
    else:
        font = st.selectbox('Шрифт', ['Times New Roman', 'Arial', 'Calibri'])
        size = st.number_input('Шрифт, пт', 10, 18, 12)
        spacing = st.select_slider('Межстрочный интервал', [1.0, 1.15, 1.5, 2.0], value=1.5)
        indent = st.number_input('Абзацный отступ, см', 0.0, 2.5, 1.25, 0.05)
        left = st.number_input('Слева, см', 0.5, 5.0, 3.0, 0.1)
        right = st.number_input('Справа, см', 0.5, 5.0, 1.5, 0.1)
        top = st.number_input('Сверху, см', 0.5, 5.0, 2.0, 0.1)
        bottom = st.number_input('Снизу, см', 0.5, 5.0, 2.0, 0.1)
        cfg = Profile(font, size, spacing, indent, left, right, top, bottom)
    st.divider()
    st.caption('Правила страницы, текста и нумерации основаны на пп. 6.1–6.3 ГОСТ 7.32–2017. Полный нормоконтроль не выполняется.')

upload, settings, about = st.tabs(['Оформить документ', 'Профиль и ограничения', 'О проекте'])
with upload:
    file = st.file_uploader('Загрузи Word-документ', type=['docx'])
    if file is not None:
        source = file.getvalue()
        if len(source) > MAX_BYTES:
            st.error('Размер DOCX превышает 15 МБ.')
            st.stop()
        source_hash = hashlib.sha256(source).hexdigest()
        try:
            doc = get_document(source)
        except Exception as error:
            st.error(f'Не получилось прочитать файл: {error}')
            st.stop()
        x, y, z = st.columns(3)
        x.metric('Абзацев', len(doc.paragraphs))
        y.metric('Таблиц', len(doc.tables))
        z.metric('Разделов Word', len(doc.sections))
        with st.expander('Проверка исходного файла', expanded=False):
            initial_issues = diagnostics(source, cfg)
            if initial_issues:
                st.dataframe(initial_issues[:150], use_container_width=True, hide_index=True)
                st.caption('Показываются первые 150 обнаруженных замечаний. Это частичный технический контроль.')
            else:
                st.success('По реализованным правилам явных замечаний не обнаружено.')
        st.subheader('Настройки обработки')
        left, right = st.columns(2)
        with left:
            add_title = st.checkbox('Добавить учебный титульный лист', value=False)
            add_toc = st.checkbox('Вставить автоматическое содержание', value=False)
        with right:
            headings = st.checkbox('Распознавать заголовки по тексту', value=True,
                                   help='Эвристика: может ошибаться. Лучше использовать стили Word «Заголовок 1–3».')
            tables = st.checkbox('Форматировать текст в таблицах', value=True)
        cover = {}
        if add_title:
            st.info('Титульный лист здесь учебный и упрощённый, а не официальный образец из приложения А ГОСТ. Для отчёта НИР потребуется отдельная форма.')
            col1, col2 = st.columns(2)
            with col1:
                cover['organization'] = st.text_input('Ведомство / организация', 'Министерство науки и высшего образования РФ')
                cover['university'] = st.text_input('Университет', 'НИТУ МИСИС')
                cover['department'] = st.text_input('Институт / кафедра', '')
                cover['work_type'] = st.text_input('Тип работы', 'ОТЧЁТ')
                cover['topic'] = st.text_input('Название работы', '')
            with col2:
                cover['student'] = st.text_input('Студент', '')
                cover['group'] = st.text_input('Группа', '')
                cover['teacher'] = st.text_input('Преподаватель', '')
                cover['city'] = st.text_input('Город', 'Москва')
                cover['year'] = st.text_input('Год', str(date.today().year))

        params = json.dumps(dict(source_hash=source_hash, cfg=cfg.summary(), add_title=add_title,
                                 add_toc=add_toc, headings=headings, tables=tables,
                                 cover=cover), ensure_ascii=False, sort_keys=True)
        job_id = hashlib.sha256(params.encode()).hexdigest()
        if st.button('✨ Оформить документ', type='primary', use_container_width=True):
            try:
                output, before, after, warnings = format_docx(
                    source, cfg, add_title=add_title, cover=cover,
                    create_toc=add_toc, recognize_headings=headings, format_tables=tables,
                )
                st.session_state['job'] = dict(id=job_id, output=output, before=before,
                                               after=after, warnings=warnings,
                                               original=file.name, profile=cfg.summary())
            except Exception as error:
                st.error(f'Ошибка обработки документа: {error}')

        job = st.session_state.get('job')
        if job and job['id'] == job_id:
            st.success('Документ обработан. Проверь результат в Microsoft Word перед сдачей.')
            output_name = file.name.removesuffix('.docx') + '_GOST.docx'
            st.download_button('⬇️ Скачать исправленный Word', job['output'], output_name,
                               'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                               use_container_width=True)
            if job['warnings']:
                with st.expander('Что проверить вручную', expanded=True):
                    for warning in dict.fromkeys(job['warnings']):
                        st.write('• ' + warning)
            st.subheader('Сравнение до и после')
            a, b = st.columns(2)
            with a:
                st.markdown('**Исходный текст**')
                st.text_area('До', job['before'][:4500], height=220, label_visibility='collapsed')
            with b:
                st.markdown('**После обработки**')
                st.text_area('После', job['after'][:4500], height=220, label_visibility='collapsed')
            changes = '\n'.join(difflib.unified_diff(job['before'].splitlines(), job['after'].splitlines(),
                                                     fromfile='До', tofile='После', lineterm=''))
            st.caption('Сравнение выше текстовое, не визуальное: изменения полей, шрифтов и интервалов в нём не отображаются.')
            with st.expander('Подробный текстовый diff'):
                st.code((changes or 'Текстовое содержимое не менялось.')[:12000], language='diff')
            issues = diagnostics(job['output'], cfg)
            with st.expander(f'Проверка результата: {len(issues)} замечаний'):
                if issues:
                    st.dataframe(issues[:200], hide_index=True, use_container_width=True)
                else:
                    st.success('По реализованным проверкам замечаний не обнаружено.')
        elif job:
            st.info('Настройки изменились. Нажми «Оформить документ» ещё раз.')
    else:
        st.markdown('Загрузи **DOCX**. Файл можно обработать и скачать без регистрации на сайте.')
        st.caption('Для конфиденциальных документов лучше использовать локальную версию на своём компьютере.')

with settings:
    st.subheader('Выбранный профиль')
    st.json(cfg.summary())
    st.write('Стандарт определяет формат A4, основные поля, полуторный интервал, чёрный шрифт не менее 12 пт и правила нумерации. Times New Roman 12 пт выбран как фиксированная настройка сервиса.')
    st.warning('Проверка не охватывает полностью формулы, библиографию, правильность ссылок, рисунки, списки, сложные таблицы, сноски и все разделы ГОСТ. Результат не является подтверждением полного соответствия.')
    st.markdown('[Текст ГОСТ 7.32–2017](https://ifap.ru/library/gost/7322017.pdf)')

with about:
    st.subheader('GOST DOC · командный проект')
    st.write('Разработчики совместно ведут исходный код через GitHub. Публичная версия развёртывается из основной ветки репозитория в Streamlit Community Cloud.')
    st.write('В публичной версии загружаемые DOCX-файлы отправляются серверу хостинга для обработки. Не загружай конфиденциальные документы без соответствующей политики и мер защиты.')
    st.markdown('**Технический стек:** Python, python-docx, Streamlit, pytest, GitHub Actions.')
