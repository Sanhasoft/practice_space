"""클라우드와 같은 조건에서의 화면 검증 (로컬 검증용, 저장소에 올리지 않아도 되는 보조 스크립트).

임시 폴더에 'git에 올라간 파일'만 복사하고(엑셀·.env·캐시 없음), API 키는 .streamlit/secrets.toml 로만 준 뒤,
화면에서 엑셀을 열고 → 사례 1건을 실제 LLM으로 추출 → 확인 → PPT 생성까지 실행한다. 키 값은 출력하지 않는다.

사용:  python scripts/verify_cloud_like.py <임시폴더>      (프로젝트 폴더에서, 설치된 파이썬으로 실행)
"""
from __future__ import annotations

import io
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

PROJECT = Path(__file__).resolve().parent.parent
MAX_TRIES = 3            # 문제 없는 사례를 찾기 위해 시도할 하루치 사례 수 (LLM 호출 = 시도 수만큼)
results: list[tuple[str, bool, str]] = []


def check(name: str, ok: bool, detail: str = '') -> bool:
    results.append((name, ok, detail))
    print(f"[{'PASS' if ok else 'FAIL'}] {name}" + (f' — {detail}' if detail else ''), flush=True)
    return ok


def build_sim(sim: Path) -> str:
    """git이 추적하는 파일만 복사하고 키를 secrets.toml로 준비한다. 반환: 키 값(출력 금지)."""
    if sim.exists():
        shutil.rmtree(sim)
    names = subprocess.run(['git', 'ls-files', '-z'], cwd=PROJECT, capture_output=True, check=True).stdout.split(b'\0')
    for raw in filter(None, names):
        name = raw.decode('utf-8')
        target = sim / name
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(PROJECT / name, target)
    key = re.search(r'UPSTAGE_API_KEY=(\S+)', (PROJECT / '.env').read_text(encoding='utf-8-sig')).group(1).strip('"\'')
    (sim / '.streamlit').mkdir(exist_ok=True)
    (sim / '.streamlit' / 'secrets.toml').write_text(f'UPSTAGE_API_KEY = "{key}"\n', encoding='utf-8')
    return key


def visible_texts(at) -> str:
    parts = []
    for kind in ('markdown', 'caption', 'info', 'success', 'error', 'warning', 'text', 'title', 'header', 'subheader'):
        parts += [getattr(el, 'value', '') or '' for el in at.get(kind)]
    return '\n'.join(parts)


def main() -> int:
    sim = Path(sys.argv[1]).resolve()
    key = build_sim(sim)
    check('임시 폴더에 엑셀·.env·캐시가 없음',
          not list(sim.glob('*.xlsx')) and not (sim / '.env').exists() and not (sim / 'extraction_cache.json').exists())
    os.chdir(sim)
    sys.path.insert(0, str(sim))
    os.environ.pop('UPSTAGE_API_KEY', None)                 # 환경변수가 아니라 Secrets로 읽는지 확인하려고 비운다

    import all_function_code as app
    from pptx import Presentation
    from streamlit.testing.v1 import AppTest

    excel = (PROJECT / 'masked_사고관리시스템_이미지포함_사고내용보완.xlsx').read_bytes()
    check('클라우드처럼 기본 엑셀이 없음', not app.XLSX.exists())

    at = AppTest.from_file(str(sim / 'app.py'), default_timeout=300).run()
    check('첫 화면: 오류 없음', not at.exception)
    check('첫 화면: 엑셀 열기 칸이 있고 조회 칸은 아직 없음', len(at.file_uploader) == 1 and len(at.selectbox) == 0)
    check('첫 화면: 엑셀을 열어 달라는 안내', any('엑셀 파일 열기' in i.value for i in at.info))

    at.file_uploader[0].set_value(('열은파일.xlsx', excel, 'application/vnd.ms-excel')).run()
    check('엑셀을 열면 조회 화면 표시', not at.exception and [s.label for s in at.selectbox] == ['사고유형', '재해정도', '사고성여부'])
    check('화면에 연 파일 이름 표시', any('열은파일.xlsx' in c.value for c in at.caption))

    df = app.build_standard_records(excel)
    per_day = df.groupby(df['날짜'].dt.date).size()
    days = [day for day, count in per_day.items() if count == 1][:MAX_TRIES]
    built = False
    for day in days:
        at.date_input(key='date_range').set_value((day, day))
        at.button[0].click().run()
        if at.exception:
            check('조회 실행', False, '예외 발생')
            break
        extracted = at.session_state['q_extracted']
        failed = [row for row, value in extracted.items() if 'error' in value]
        if failed:
            check('LLM 추출(Secrets 키)', False, f'{day}: 추출 실패 {failed}')
            continue
        check('LLM 추출(Secrets 키로 실제 호출)', True, f'{day} 1건')
        cache = sim / 'extraction_cache.json'
        check('추출 결과가 서버 캐시에 저장됨(실제 호출의 증거)', cache.exists() and cache.stat().st_size > 10)
        version = at.session_state['q_version']
        table = app.build_review_table(at.session_state['q_result']['rows'], extracted)
        need = [i for i, differs in enumerate(table['규칙≠LLM']) if differs]
        at.session_state[f'editor_{version}'] = {'edited_rows': {i: {'장소 승인': True} for i in need},
                                                 'added_rows': [], 'deleted_rows': []}
        button = next(b for b in at.button if b.label == '예, PPT 만들기')
        if button.disabled:
            check('PPT 만들기 가능(사례에 해결할 문제 없음)', False, f'{day}: 문제 있는 사례라 다음 날짜 시도')
            continue
        button.click().run()
        ppt = at.session_state['q_ppt']['data'] if 'q_ppt' in at.session_state else None
        slides = len(Presentation(io.BytesIO(ppt)).slides) if ppt else 0
        check('PPT 생성(표지·목차·구분 + 사례 1장 = 4장)', not at.exception and slides == 4, f'{slides}장')
        check('다운로드 버튼 표시', len(at.get('download_button')) == 1)
        built = True
        break
    check('조회→추출→PPT 전체 흐름 완료', built)

    shown = visible_texts(at)
    check('화면 어디에도 API 키가 보이지 않음', key not in shown)
    leaked = [str(p.relative_to(sim)) for p in sim.rglob('*')
              if p.is_file() and p.name != 'secrets.toml' and '.git' not in p.parts and key.encode() in p.read_bytes()]
    check('임시 폴더의 다른 파일(캐시 등)에도 API 키가 없음', not leaked, ', '.join(leaked))

    failed = [name for name, ok, _ in results if not ok]
    print(f'\n검증 {len(results) - len(failed)}/{len(results)} 통과' + (f' — 실패: {failed}' if failed else ''))
    return 1 if failed else 0


if __name__ == '__main__':
    sys.exit(main())
