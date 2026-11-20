"""Technical table checks, distinct from manual semantic/pagination checks."""
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Length

from .document import iter_cells, iter_runs, iter_tables, paragraph_value, run_value
from .formatter import is_list, table_neighbors
from .misis import is_equation_table


def audit_tables(doc, cfg):
    issues = []
    category = 'Таблицы'

    def add(level, location, text):
        issues.append({'Уровень': level, 'Категория': category,
                       'Место': location, 'Замечание': text})

    def check_paragraph(p, location, expected, *, before=0, alignment=WD_ALIGN_PARAGRAPH.LEFT):
        if is_list(p):
            add('Проверить вручную', location, 'В ячейке сохранён список; проверьте отступы и маркеры Word.')
        if paragraph_value(p, 'alignment', doc) != alignment:
            add('Ошибка', location, f'Выравнивание текста: требуется {alignment.name}.')
        for attr, label in [('first_line_indent', 'Красная строка'),
                            ('left_indent', 'Отступ слева'), ('right_indent', 'Отступ справа')]:
            value = paragraph_value(p, attr, doc)
            if abs(value.cm) > .03:
                add('Ошибка', location, f'{label}: {value.cm:.2f} см; требуется 0 см.')
        for attr, label, target in [('space_before', 'перед', before), ('space_after', 'после', 0)]:
            value = paragraph_value(p, attr, doc)
            if abs(value.pt - target) > .1:
                add('Ошибка', location, f'Интервал {label}: {value.pt:g} пт; требуется {target:g} пт.')
        spacing = paragraph_value(p, 'line_spacing', doc)
        if isinstance(spacing, Length) or abs(spacing - expected.line_spacing) > .03:
            actual = f'{spacing.pt:g} пт (фиксированный)' if isinstance(spacing, Length) else f'{spacing:g}'
            add('Ошибка', location, f'Межстрочный интервал {actual}; требуется {expected.line_spacing:g}.')
        names, sizes = set(), set()
        for run in iter_runs(p):
            if run.text.strip():
                names.add(run_value(run, p, 'font', doc))
                sizes.add(run_value(run, p, 'size', doc))
        if None in names or None in sizes:
            add('Проверить вручную', location, 'Не удалось однозначно определить унаследованный шрифт/размер.')
        wrong_fonts = sorted(name for name in names if name is not None and name != expected.font)
        if wrong_fonts:
            add('Ошибка', location, f'Шрифт: {", ".join(wrong_fonts)}; требуется {expected.font}.')
        wrong_sizes = sorted(size for size in sizes if size is not None and abs(size - expected.size) > .2)
        if wrong_sizes:
            add('Ошибка', location, f'Размер: {", ".join(f"{size:g}" for size in wrong_sizes)} пт; требуется {expected.size:g} пт.')

    for location, table in iter_tables(doc):
        if cfg.ruleset == 'misis':
            from .cover import cover_elements, inside_cover
            if inside_cover(table._tbl, set(cover_elements(doc))):
                add('Проверить вручную', location, 'Таблица принадлежит титульному листу; сравните его с эталоном (стр. 9).')
                continue
        category = 'Формулы' if cfg.ruleset == 'misis' and is_equation_table(table) else 'Таблицы'
        if category == 'Формулы':
            for column, cell in enumerate(table.rows[0].cells):
                for index, p in enumerate(cell.paragraphs, 1):
                    check_paragraph(p, f'{location}, формула, ячейка {column + 1}, абзац {index}', cfg,
                                    alignment=(WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT)[column])
            add('Проверить вручную', location, 'Проверьте смысл формулы, пояснения «где», нумерацию и отсутствие границ (стр. 10).')
            continue
        if cfg.ruleset == 'misis':
            if table.alignment != WD_TABLE_ALIGNMENT.CENTER:
                add('Ошибка', location, 'Таблица должна быть расположена по центру.')
            width = table._tbl.tblPr.find(qn('w:tblW'))
            if width is None or width.get(qn('w:type')) != 'pct' or width.get(qn('w:w')) != '5000':
                add('Ошибка', location, 'Предпочитаемая ширина таблицы должна быть 100% (тип «Проценты»).')
        for row, col, cell in iter_cells(table):
            cell_location = f'{location}, строка {row}, ячейка {col}'
            for index, paragraph in enumerate(cell.paragraphs, 1):
                check_paragraph(paragraph, f'{cell_location}, абзац {index}', cfg.for_table())
            if cfg.ruleset == 'misis':
                shading = cell._tc.xpath('./w:tcPr/w:shd')
                if shading:
                    node = shading[0]
                    fill = node.get(qn('w:fill'), 'auto').lower()
                    pattern = node.get(qn('w:val'), 'clear')
                    theme = any(key in node.attrib for key in [qn('w:themeFill'), qn('w:themeColor')])
                    if fill not in ('auto', '') or pattern not in ('clear', 'nil') or theme:
                        add('Ошибка', cell_location, 'В таблице не допускается заливка ячеек.')
                elif table.style and table.style.element.xpath('.//w:shd'):
                    add('Проверить вручную', cell_location, 'Возможна заливка из стиля таблицы; проверьте в Word.')
        if cfg.ruleset == 'misis':
            caption, after = table_neighbors(table)
            in_body = table._tbl.getparent().tag == qn('w:body')
            if caption is not None:
                check_paragraph(caption, f'{location}, подпись', cfg)
            elif in_body:
                add('Проверить вручную', location, 'Проверьте название таблицы: рядом не распознана текстовая подпись.')
            if after is not None:
                before = paragraph_value(after, 'space_before', doc)
                if abs(before.pt - 6) > .1:
                    add('Ошибка', f'{location}, текст после таблицы', 'Перед текстом после таблицы требуется интервал 6 пт.')
            elif in_body:
                add('Проверить вручную', location, 'После таблицы нужен текст/вывод; не завершайте раздел таблицей.')
            if table.rows:
                first = table.rows[0]
                if not first._tr.xpath('./w:trPr/w:tblHeader'):
                    add('Проверить вручную', location, 'Подтвердите наличие смысловой шапки и её повтор на следующих страницах.')
                if first.cells and first.cells[0].text.strip().lower() in ('№', '№ п/п', 'п/п', 'номер'):
                    add('Проверить вручную', location, 'Возможна колонка с порядковым номером; методичка её запрещает. Данные не удаляются автоматически.')
    return issues
