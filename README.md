# 사고사례 PPT 생성기

사고관리 엑셀에서 사고 케이스를 조회하고, LLM(Upstage Solar)으로 제목·장소·내용을 추출한 뒤,
확인·수정을 거쳐 안전사고사례 PPT를 만드는 Streamlit 앱입니다.

## 사용 방법

1. 앱에 접속해 화면 맨 위 **엑셀 파일 열기** 칸에서 사고관리 엑셀(.xlsx)을 선택(업로드)합니다. 다른 엑셀을 고르면 조회 조건이 새 파일 기준으로 초기화됩니다.
2. 사고 일자와 사고유형·재해정도·사고성여부를 고르고 **선택한 조건으로 조회**를 누릅니다.
3. 추출된 제목·장소·소속·내용을 표에서 확인·수정하고, 문제가 없으면 **PPT 만들기 → 내려받기**를 누릅니다.

### 데이터 취급 안내

- 업로드한 엑셀은 그 접속(세션) 동안 서버 메모리에서만 쓰이고 **저장소에는 저장되지 않습니다.**
- 조회한 사고 내용 텍스트는 제목·장소·내용 추출을 위해 **Upstage(외부 LLM 서비스)로 전송**됩니다.
- 추출 결과(요약)는 앱 서버의 캐시 파일에 임시로 남고, 앱이 재시작되면 사라집니다.
- 이 앱은 링크를 아는 누구나 쓸 수 있고 **모든 접속이 하나의 API 키를 함께 씁니다.** 추출 건수 상한은 없으므로
  Upstage 콘솔에서 사용량을 주기적으로 확인하세요.

## 내 PC에서 실행

```
pip install -r requirements.txt      # Python 3.11 이상
python -m streamlit run app.py
```

- 프로젝트 폴더에 `masked_사고관리시스템_이미지포함_사고내용보완.xlsx`가 있으면 업로드 없이 그 파일을 씁니다.
- LLM 추출에는 API 키가 필요합니다. `.env.example`을 `.env`로 복사해 키를 채우거나
  `.streamlit/secrets.toml.example`을 `.streamlit/secrets.toml`로 복사해 채웁니다. 키가 없으면 조회만 되고 LLM 추출은 꺼집니다.

## API 키 관리 원칙

- 키는 **저장소에 절대 올리지 않습니다.** `.env`, `.streamlit/secrets.toml`은 `.gitignore` 대상입니다.
- Streamlit Cloud에서는 앱 설정의 **Secrets** 칸에만 `UPSTAGE_API_KEY = "..."`를 넣습니다. Secrets를 바꾸면 앱이 재시작됩니다.
- 커밋 전 검사를 켭니다(클론한 뒤 한 번만):
  ```
  git config core.hooksPath .githooks
  ```
  `scripts/check_no_secrets.py`가 커밋마다 실행되어 키 모양 문자열, `.env`·엑셀·캐시·생성 PPT 같은 파일이 올라가는지 검사하고
  발견하면 커밋을 막습니다. 검사 결과에는 파일 경로만 나오고 키 값은 출력하지 않습니다.
- 키가 노출된 것으로 의심되면 Upstage 콘솔에서 폐기·재발급하고 Secrets(과 로컬 `.env`) 값을 바꿉니다.
  저장소 이력에서 지우는 것보다 키를 폐기하는 것이 먼저입니다.

## 구성

| 파일 | 역할 |
|---|---|
| `app.py` | Streamlit 진입점 (`streamlit run app.py`) |
| `all_function_code.py` | 엑셀 변환·조회·LLM 추출·PPT 생성·화면 전체 |
| `(ref1)AIglue_안전사고사례_PPT마스터슬라이드.pptx` | PPT 아이콘 출처 템플릿 |
| `scripts/check_no_secrets.py`, `.githooks/pre-commit` | 커밋 전 비밀·데이터 유출 검사 |
| `scripts/verify_cloud_like.py` | 클라우드와 같은 조건(엑셀·`.env` 없음, Secrets 키)에서 화면 흐름을 검증하는 보조 스크립트 |
| `.streamlit/config.toml` | 화면 오류 상세 숨김, 업로드 크기 상한 등 |
| `requirements.txt` | 배포 시 설치할 라이브러리 (테스트한 버전으로 고정) |
