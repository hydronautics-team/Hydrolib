"""Human-readable Actions result. Only reads artifacts, never executes PR data."""
import html
import json
import os
from pathlib import Path


def render(directory, outcome, repo, number, fork=False):
    directory = Path(directory)
    context_path = directory / 'context.json'
    if context_path.exists():
        context = json.loads(context_path.read_text(encoding='utf-8'))
        fork, number = context['fork'], context['number']
    report_path = directory / 'report.json'
    report = json.loads(report_path.read_text(encoding='utf-8')) if report_path.exists() else {}
    conflicts = report.get('conflicts', [])
    unresolved = [c for c in conflicts if c.get('choice') is None]
    esc = lambda value: html.escape(str(value))
    url = 'https://github.com/' + repo + '/actions/workflows/'
    result_path = directory / 'result.json'
    result = json.loads(result_path.read_text(encoding='utf-8')) if result_path.exists() else {}
    if outcome == 'success':
        title = '✅ Книгу удалось объединить'
        message = 'Проверка успешна. Автоматическая проверка не меняет ветки. Перед использованием результата откройте книгу в Excel, пересчитайте и сохраните.'
    elif unresolved:
        title = '⚠️ Не удалось объединить книгу: нужен выбор стороны'
        message = f'Неразрешённых конфликтов: {len(unresolved)}. Это расхождения между версиями, а не ошибка запуска.'
    else:
        title = '❌ Проверка не завершилась'
        error = directory / 'failure.txt'
        message = report.get('message') or (error.read_text(encoding='utf-8') if error.exists() else 'Проверка книги не запускалась или остановилась раньше. Откройте красный шаг в разделе Jobs.')
    if outcome == 'success' and result:
        message = 'Объединение подготовлено. Перед использованием пересчитайте книгу в Excel и проверьте diff результата.'
    parts = [f'<h2>{title}</h2><p>{esc(message)}</p>', f'<p>PR: <a href="https://github.com/{esc(repo)}/pull/{esc(number)}">#{esc(number)}</a>. Наше (ours) — ветка PR; их (theirs) — целевая ветка.</p>']
    def display(value):
        if value == {'missing': True}:
            return 'Удалено / отсутствует'
        if isinstance(value, dict) and 'sha256' in value:
            return 'Бинарная версия ' + value['sha256'][:12]
        text = value if isinstance(value, str) else json.dumps(value, ensure_ascii=False)
        return esc(text[:500])
    for heading, records in [('Нерешённые конфликты', unresolved),
                             ('Разрешённые конфликты', [c for c in conflicts if c.get('choice') is not None])]:
        if not records:
            continue
        parts.append(f'<h3>{heading}: {len(records)}</h3><table><tr><th>Место</th><th>Исходное</th><th>Наше</th><th>Их</th><th>Решение</th></tr>')
        for c in records:
            place = ' / '.join(str(c[k]) for k in ('sheet', 'part', 'field') if k in c)
            if c.get('kind') == 'binary_part':
                place = 'VBA: макросы различаются (выбираются целиком)'
            elif c.get('kind') == 'column_layout':
                place += ' / Оформление колонок'
            elif c.get('kind') == 'record':
                place += ' / Запись целиком (добавление, изменение или удаление)'
            parts.append(f'<tr><td>{esc(place)}</td><td>{display(c.get("base"))}</td><td>{display(c.get("ours"))}</td><td>{display(c.get("theirs"))}</td><td>{esc(c.get("choice") or "Не выбрано")}</td></tr>')
        parts.append('</table>')
    if outcome != 'success' and not unresolved:
        parts.append('<h3>Что мешает завершить слияние</h3><p>' + esc(message) + '</p>')
    parts.append('<h3>Что делать дальше</h3>')
    if result.get('mode') == 'fork_result':
        parts.append(f'<p>PR с результатом: <a href="{esc(result["url"])}">#{esc(result["number"])}</a>. Исходный PR и личный форк не изменены. Проверьте и примите новый PR; затем исходный можно закрыть вручную.</p>')
    elif result.get('mode') == 'updated':
        parts.append('<p>Объединение записано в ветку исходного PR.</p>')
    elif fork and (unresolved or outcome == 'success'):
        parts.append(f'<p>PR из форка. Откройте <a href="{url}xlsm-update.yml">Объединить XLSM</a> → Run workflow → номер PR {esc(number)} → ours/theirs. Результат будет в новой ветке и новом PR общего репозитория. Личный форк и исходный PR не изменятся.</p>')
    elif unresolved:
        parts.append(f'<p>Откройте <a href="{url}xlsm-update.yml">Объединить XLSM</a> → Run workflow → master → номер PR {esc(number)} → ours или theirs → Run workflow. Независимые изменения сохранятся. При расхождении VBA выбирается весь VBA-проект стороны.</p>')
    elif outcome != 'success':
        parts.append('<p>Если причина — структура книги, шрифты или другие неподдерживаемые изменения, нужен ручной разбор. Выбор ours/theirs не обходит эти ограничения.</p>')
    parts.append(f'<p><a href="{url}xlsm-check.yml">Повторить проверку по номеру PR</a>. Подробности и исходные книги находятся в Artifacts этого запуска.</p>')
    return '\n'.join(parts)


if __name__ == '__main__':
    output = render('xlsm-artifacts', os.environ.get('MERGE_OUTCOME', 'failure'),
                    os.environ['GITHUB_REPOSITORY'], os.environ.get('PR_NUMBER', ''),
                    os.environ.get('IS_FORK') == 'true')
    with open(os.environ['GITHUB_STEP_SUMMARY'], 'a', encoding='utf-8') as stream:
        stream.write(output)

