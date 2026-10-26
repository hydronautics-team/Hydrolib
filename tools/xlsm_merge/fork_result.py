"""Publish a resolved fork snapshot to the target repository, never to the fork."""
import hashlib
import json
from io import BytesIO
from pathlib import Path
from urllib.parse import urlencode
from zipfile import ZipFile


def snapshot(pr):
    return (pr['state'], pr['head']['sha'], pr['base']['sha'], pr['head']['ref'],
            pr['base']['ref'], pr['head']['repo']['full_name'])


def check_snapshot(api, number, original):
    if snapshot(api('pulls/' + number)) != snapshot(original):
        raise ValueError('PR изменился во время объединения. Запустите workflow заново; исходная ветка не изменена.')


def workbook_parts(data):
    with ZipFile(BytesIO(data)) as z:
        return {name: z.read(name) for name in z.namelist()}


def publish(git, api, repo, number, pr, worktree, workbook, policy, resolutions, artifacts):
    # A deterministic branch per input+decision+tool snapshot allows retry after
    # API errors without force push and without duplicate replacement PRs.
    tool = Path(__file__).with_name('merge.py').read_bytes()
    identity = json.dumps({'head': pr['head']['sha'], 'base': pr['base']['sha'],
        'target': pr['base']['ref'], 'policy': policy, 'resolutions': resolutions,
        'tool': hashlib.sha256(tool).hexdigest()}, sort_keys=True).encode()
    key = hashlib.sha256(identity).hexdigest()[:24]
    branch = f'xlsm-resolved/pr-{number}-{key}'
    ref = 'refs/heads/' + branch
    check_snapshot(api, number, pr)
    remote = git('ls-remote', '--heads', 'origin', ref, capture_output=True, text=True).stdout.strip()
    if remote:
        existing = remote.split()[0]
        git('fetch', 'origin', ref)
        # Never replace an existing branch, even one with the expected name.
        git('merge-base', '--is-ancestor', pr['head']['sha'], existing)
        git('merge-base', '--is-ancestor', pr['base']['sha'], existing)
        git('diff', '--exit-code', existing, 'HEAD', '--', '.', ':!' + workbook, cwd=worktree)
        a = git('show', existing + ':' + workbook, capture_output=True).stdout
        b = git('show', 'HEAD:' + workbook, cwd=worktree, capture_output=True).stdout
        if workbook_parts(a) != workbook_parts(b):
            raise ValueError('Ветка результата уже существует с другим содержимым; перезапись запрещена.')
    else:
        git('push', 'origin', 'HEAD:' + ref, cwd=worktree)
    check_snapshot(api, number, pr)
    query = urlencode({'state': 'all', 'head': repo.split('/')[0] + ':' + branch,
                       'base': pr['base']['ref']})
    found = api('pulls?' + query)
    if found:
        replacement = found[0]
    else:
        replacement = api('pulls', method='POST', payload={
            'title': f'Объединение XLSM из PR #{number}',
            'head': branch, 'base': pr['base']['ref'],
            'body': f'Результат объединения исходного PR #{number}. Исходный PR и форк не изменены.\n\n'
                    f'Источник: {pr["head"]["repo"]["full_name"]}, commit {pr["head"]["sha"]}.\n'
                    f'Целевая версия: {pr["base"]["sha"]}. Выбор конфликтов: {policy}.\n\n'
                    'Ветка включает все изменения исходного PR, а не только XLSM. '
                    'Проверьте diff. Перед использованием книги пересчитайте её в Excel и сохраните. '
                    'Если различался VBA-проект, он выбран целиком. После принятия результата '
                    'исходный PR можно закрыть вручную как заменённый этим PR.'})
    result = {'mode': 'fork_result', 'branch': branch, 'url': replacement['html_url'],
              'number': replacement['number'], 'source_pr': number}
    Path(artifacts, 'result.json').write_text(json.dumps(result, ensure_ascii=False), encoding='utf-8')
    return result
