"""커밋 전 검사: API 키 같은 비밀과 데이터 파일이 저장소에 올라가지 않게 막는다.

사용:  python scripts/check_no_secrets.py          스테이징된(커밋하려는) 파일 검사, 문제가 있으면 종료 코드 1
       python scripts/check_no_secrets.py --all    git이 추적 중인 모든 파일 검사
설치:  git config core.hooksPath .githooks         (커밋할 때마다 자동 실행)
결과에는 파일 경로와 이유만 출력하고 비밀 값은 절대 출력하지 않는다.
"""
from __future__ import annotations

import re
import subprocess
import sys
import tomllib
from collections.abc import Mapping, Sequence
from pathlib import Path, PurePosixPath

KEY_SHAPED = re.compile(rb'(?<![A-Za-z0-9])up_[A-Za-z0-9]{16,}')   # Upstage 키 모양 ('_'는 경계로 보지 않는다)
SECRET_NAME_WORDS = ('KEY', 'TOKEN', 'SECRET', 'PASSWORD')          # 비밀로 보는 항목 이름
MIN_SECRET_LENGTH = 8                                                # 이보다 짧은 값은 우연히 겹치기 쉬워 검색에 쓰지 않는다
ALLOWED_FILES = {'(ref1)aiglue_안전사고사례_ppt마스터슬라이드.pptx'}         # 올려도 되는 이진 파일(템플릿)
DATA_SUFFIXES = {'.xlsx': '엑셀 데이터', '.xls': '엑셀 데이터', '.xlsm': '엑셀 데이터', '.xlsb': '엑셀 데이터',
                 '.csv': '표 데이터', '.parquet': '표준화 데이터', '.tmp': '임시 파일(캐시 조각일 수 있음)',
                 '.ipynb': '노트북(작업 기록·출력에 내부 내용이 있을 수 있음)',
                 '.pptx': '생성된 PPT(사고 데이터가 들어 있을 수 있음)', '.pdf': '문서(내부 자료일 수 있음)',
                 '.pem': '개인 키 파일', '.key': '개인 키 파일'}


def _blocked_reason(path: str) -> str | None:
    name = PurePosixPath(path.replace('\\', '/')).name.lower()
    if name in ALLOWED_FILES:
        return None
    if name == '.env' or (name.startswith('.env.') and name != '.env.example'):
        return '환경 파일(.env)에는 API 키가 들어 있습니다'
    if name == 'secrets.toml':
        return 'Streamlit Secrets 파일에는 API 키가 들어 있습니다'
    if name == 'extraction_cache.json':
        return '사고 내용을 요약한 캐시입니다'
    return DATA_SUFFIXES.get(PurePosixPath(name).suffix)


def find_problems(files: Mapping[str, bytes], secret_values: Sequence[str]) -> list[str]:
    """파일 이름과 내용을 검사해 문제 목록(경로와 이유만)을 돌려준다. 비밀 값은 결과에 넣지 않는다."""
    needles = [value.strip().encode('utf-8') for value in secret_values if len(value.strip()) >= MIN_SECRET_LENGTH]
    problems = []
    for path, content in files.items():
        reason = _blocked_reason(path)
        if reason:
            problems.append(f'{path}: 올리면 안 되는 파일입니다 ({reason})')
        if KEY_SHAPED.search(content):
            problems.append(f'{path}: API 키 모양의 문자열이 들어 있습니다')
        elif any(needle in content for needle in needles):
            problems.append(f'{path}: 로컬 비밀(.env·secrets.toml)의 값과 같은 문자열이 들어 있습니다')
    return problems


def _is_secret_name(name: str) -> bool:
    return any(word in name.upper() for word in SECRET_NAME_WORDS)


def local_secret_values(env_path: Path) -> list[str]:
    """.env에서 이름에 KEY·TOKEN·SECRET·PASSWORD가 들어간 항목의 값 목록 (없으면 빈 목록)"""
    if not env_path.is_file():
        return []
    values = []
    for line in env_path.read_text(encoding='utf-8-sig').splitlines():
        name, _, value = line.strip().partition('=')
        value = value.split(' #')[0].strip().strip('"').strip("'")     # 줄 끝 주석은 값이 아니다
        if value and _is_secret_name(name):
            values.append(value)
    return values


def secrets_toml_values(path: Path) -> list[str]:
    """Streamlit secrets.toml에서 이름에 KEY·TOKEN·SECRET·PASSWORD가 들어간 문자열 값 목록 (없거나 깨졌으면 빈 목록)"""
    if not path.is_file():
        return []
    try:
        data = tomllib.loads(path.read_text(encoding='utf-8-sig'))
    except (OSError, ValueError):
        return []
    values: list[str] = []

    def walk(node: object, name: str = '') -> None:
        if isinstance(node, dict):
            for key, child in node.items():
                walk(child, str(key))
        elif isinstance(node, str) and node.strip() and _is_secret_name(name):
            values.append(node)

    walk(data)
    return values


def _git(root: Path, *args: str) -> bytes:
    return subprocess.run(['git', *args], cwd=root, capture_output=True, check=True).stdout


def _names(output: bytes) -> list[str]:
    return [name.decode('utf-8') for name in output.split(b'\0') if name]


def staged_files(root: Path) -> dict[str, bytes]:
    """커밋하려고 스테이징된 파일의 (경로 → 스테이징된 내용)"""
    names = _names(_git(root, 'diff', '--cached', '--name-only', '--diff-filter=ACMR', '-z'))
    return {name: _git(root, 'show', f':{name}') for name in names}


def tracked_files(root: Path) -> dict[str, bytes]:
    """git이 추적 중인 모든 파일의 (경로 → 작업 폴더의 내용)"""
    return {name: (root / name).read_bytes() for name in _names(_git(root, 'ls-files', '-z')) if (root / name).is_file()}


def main(argv: Sequence[str] | None = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    try:
        root = Path(_git(Path.cwd(), 'rev-parse', '--show-toplevel').decode('utf-8').strip())
        files = tracked_files(root) if '--all' in args else staged_files(root)
    except (OSError, subprocess.CalledProcessError) as error:       # 검사를 못 하면 안전하게 막는다
        print(f'git 정보를 읽지 못해 검사를 완료하지 못했습니다({type(error).__name__}). 커밋을 막습니다.')
        return 1
    secret_values = local_secret_values(root / '.env') + secrets_toml_values(root / '.streamlit' / 'secrets.toml')
    problems = find_problems(files, secret_values)
    if problems:
        print('커밋을 막았습니다. 아래 문제를 해결한 뒤 다시 시도하세요.')
        for problem in problems:
            print(f'  - {problem}')
        return 1
    print(f'비밀·데이터 검사 통과 ({len(files)}개 파일)')
    return 0


if __name__ == '__main__':
    sys.exit(main())
