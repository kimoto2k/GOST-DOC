"""Smoke test: Streamlit should render its main interface without exceptions."""
from pathlib import Path

from streamlit.testing.v1 import AppTest


def test_app_renders_homepage():
    app_path = Path(__file__).resolve().parents[1] / "app.py"
    at = AppTest.from_file(str(app_path), default_timeout=25).run()
    assert not at.exception, [str(exc) for exc in at.exception]
    assert at.sidebar.radio[0].value == 'МИСИС · методичка 2024'
    assert any('Оформление отчёта' in item.value for item in at.markdown)
    assert any('10 пт' in item.value for item in at.sidebar.caption)


def test_check_only_mode_never_offers_format_action():
    app_path = Path(__file__).resolve().parents[1] / 'app.py'
    at = AppTest.from_file(str(app_path), default_timeout=25).run()
    next(r for r in at.radio if r.label == 'Режим работы').set_value('Только проверить').run()
    assert not at.exception
    assert any(b.label == 'Проверить оформление' and b.disabled for b in at.button)
    assert not any(b.label == 'Оформить документ' for b in at.button)
    assert all(c.disabled for c in at.checkbox)


def test_custom_body_and_table_parameters_are_independent():
    app_path = Path(__file__).resolve().parents[1] / 'app.py'
    at = AppTest.from_file(str(app_path), default_timeout=25).run()
    at.sidebar.radio[0].set_value('Свой профиль').run()
    assert not at.exception
    values = {item.label: item for item in at.sidebar.number_input}
    assert values['Размер текста, пт'].value == 12
    assert values['Размер текста таблиц, пт'].value == 10
    values['Размер текста, пт'].set_value(14.0).run()
    assert not at.exception
    values = {item.label: item.value for item in at.sidebar.number_input}
    assert values['Размер текста, пт'] == 14
    assert values['Размер текста таблиц, пт'] == 10
    at.sidebar.radio[0].set_value('ГОСТ 7.32–2017').run()
    assert not at.exception
    assert any('Таблицы: 12 пт' in item.value for item in at.sidebar.caption)
