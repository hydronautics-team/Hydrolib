"""Trusted CI entrypoint: inspect PR data without executing code from the PR."""
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import urllib.request

ROOT = Path(__file__).resolve().parents[2]
WORKBOOK = 'Database/Hydrolib_bd.xlsm'
ARTIFACTS = ROOT / 'xlsm-artifacts'
ARTIFACTS.mkdir(exist_ok=True)


def git(*args, cwd=None, **kwargs):
    return subprocess.run(['git', *args], cwd=cwd or ROOT, check=True, **kwargs)


def main():
    number = os.environ['PR_NUMBER']
    if not number.isdecimal():
        raise ValueError('PR_NUMBER must be numeric')
    repo = os.environ['GITHUB_REPOSITORY']
    req = urllib.request.Request(f'https://api.github.com/repos/{repo}/pulls/{number}', headers={
        'Authorization': 'Bearer ' + os.environ['GH_TOKEN'], 'Accept': 'application/vnd.github+json'})
    with urllib.request.urlopen(req) as response:
        pr = json.load(response)
    if pr['state'] != 'open':
        raise ValueError('PR must be open')
    conflict_side = os.environ.get('CONFLICT_SIDE', 'manual')
    if conflict_side not in ('manual', 'ours', 'theirs'):
        raise ValueError('CONFLICT_SIDE must be manual, ours or theirs')
    apply = '--apply' in sys.argv
    if apply and pr['head']['repo']['full_name'] != repo:
        raise ValueError('Apply is supported only for branches in this repository; forks use local merge')
    git('config', 'core.hooksPath', '/dev/null')
    git('fetch', 'origin', pr['base']['sha'], pr['head']['sha'])
    base_sha = pr['base']['sha']
    head_sha = pr['head']['sha']
    ancestor = git('merge-base', base_sha, head_sha, capture_output=True, text=True).stdout.strip()
    with tempfile.TemporaryDirectory() as temp:
        resolution_path = None
        resolution_json = os.environ.get('RESOLUTIONS_JSON', '').strip()
        if resolution_json:
            if len(resolution_json.encode()) > 1_000_000:
                raise ValueError('Resolution JSON is too large')
            json.loads(resolution_json)  # Data only; never executed as code.
            resolution_path = Path(temp) / 'resolutions.json'
            resolution_path.write_text(resolution_json, encoding='utf-8')
        paths = []
        for label, sha in [('ancestor', ancestor), ('ours', head_sha), ('theirs', base_sha)]:
            p = ARTIFACTS / (label + '.xlsm')
            p.write_bytes(git('show', f'{sha}:{WORKBOOK}', capture_output=True).stdout)
            paths.append(str(p))
        command = [sys.executable, str(ROOT / 'tools/xlsm_merge/merge.py'), *paths,
                   str(ARTIFACTS / 'Hydrolib_bd.xlsm'), '--conflict-side', conflict_side, '--report', str(ARTIFACTS / 'report.json')]
        if resolution_path:
            command += ['--resolutions', str(resolution_path)]
        subprocess.run(command, check=True)
        if apply:
            # Use the trusted driver's absolute path, even if the PR edits tools/.
            subprocess.run([sys.executable, str(ROOT / 'tools/xlsm_merge/install.py')], cwd=ROOT, check=True)
            driver = git('config', '--get', 'merge.hydrolib-xlsm.driver', capture_output=True, text=True).stdout.strip()
            driver += ' --conflict-side ' + conflict_side
            if resolution_path:
                driver += ' --resolutions "' + str(resolution_path) + '"'
            git('config', 'merge.hydrolib-xlsm.driver', driver + ' --report "' + str(ARTIFACTS / 'git-report.json') + '"')
            info = Path(git('rev-parse', '--git-path', 'info/attributes', capture_output=True, text=True).stdout.strip())
            if not info.is_absolute():
                info = ROOT / info
            info.parent.mkdir(exist_ok=True, parents=True)
            info.write_text(WORKBOOK + ' merge=hydrolib-xlsm\n')
            worktree = Path(temp) / 'merge'
            git('worktree', 'add', '--detach', str(worktree), head_sha)
            try:
                git('config', 'user.name', 'github-actions[bot]')
                git('config', 'user.email', '41898282+github-actions[bot]@users.noreply.github.com')
                git('merge', '--no-edit', base_sha, cwd=worktree)
                branch = pr['head']['ref']
                git('check-ref-format', 'refs/heads/' + branch)
                git('push', 'origin', 'HEAD:refs/heads/' + branch, cwd=worktree)
            finally:
                git('worktree', 'remove', '--force', str(worktree))
    print('Workbook merged. Open in Excel and recalculate before use in Altium.')


if __name__ == '__main__':
    try:
        main()
    except Exception as error:
        (ARTIFACTS / 'failure.txt').write_text(str(error), encoding='utf-8')
        print(error, file=sys.stderr)
        sys.exit(1)
