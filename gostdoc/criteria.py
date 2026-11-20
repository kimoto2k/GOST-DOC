"""Single versioned source of MISIS parameters, transcribed from the supplied PDF."""
import json
from pathlib import Path

CATALOG = json.loads(Path(__file__).with_name('misis_2024.json').read_text(encoding='utf-8'))
MISIS_PARAMETERS = CATALOG['profile']


def criteria_rows():
    return [{'Критерий': rule['id'], 'Страницы PDF': ', '.join(map(str, rule['pages'])),
             'Требование': rule['text'], 'Как проверяется': rule['control']}
            for rule in CATALOG['criteria']]
