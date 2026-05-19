# Web DOCX Linker

영문 웹페이지에 걸린 하이퍼링크를 추출한 뒤, 이미 한국어로 번역된 Word 문서(`.docx`)의 대응 문장/단어 위치에 같은 링크를 넣는 도구입니다.

## 준비

```bash
cd /Users/thoamas-home/Documents/Codex/web_docx_linker
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
export OPENAI_API_KEY="여기에_API_키"
```

## 실행

```bash
python link_webpage_to_docx.py \
  --url "https://example.com/original-page" \
  --docx "/path/to/translated.docx" \
  --out "/path/to/translated_with_links.docx"
```

## 웹 UI 실행

```bash
./start_web_ui.sh
```

브라우저에서 `http://127.0.0.1:8765`로 접속한 뒤, 영문 웹페이지 주소와 번역된 `.docx` 파일을 넣으면 됩니다.

종료할 때는 아래를 실행합니다.

```bash
./stop_web_ui.sh
```

웹 화면에서 할 수 있는 일:

- 영문 웹페이지 URL 입력
- 번역된 Word 문서 업로드
- OpenAI API 키 최초 1회 입력 후 자동 저장
- 모델, 최소 신뢰도, 최대 링크 수 설정
- 링크가 삽입된 Word 파일 다운로드
- 매칭 리포트 CSV 다운로드

API 키는 `web_docx_linker/.openai_api_key`에 저장됩니다. 웹 화면의 `키 다시 입력`을 누르면 저장된 키를 지우고 새로 입력할 수 있습니다.

공개 서버 배포 방법은 `DEPLOY.md`를 참고하세요. 공개 서버 모드에서는 API 키를 서버에 저장하지 않고, 사용자 브라우저에만 저장합니다.

## 동작 방식

1. 웹페이지 HTML에서 `<a href="...">링크 텍스트</a>`를 추출합니다.
2. Word 문서의 문단 텍스트를 읽습니다.
3. OpenAI 모델이 원문 링크 텍스트와 주변 문맥을 보고, 번역 문서의 어느 문단/문구에 링크를 넣을지 판단합니다.
4. 판단된 한국어 문구가 문서에 실제로 존재하면 해당 범위에 Word 하이퍼링크를 삽입합니다.

## 옵션

- `--model`: 사용할 OpenAI 모델입니다. 기본값은 `gpt-4.1-mini`입니다.
- `--min-confidence`: 모델 매칭 신뢰도 기준입니다. 기본값은 `0.62`입니다.
- `--max-links`: 테스트용으로 처리할 링크 수를 제한합니다.
- `--dry-run`: Word 파일을 수정하지 않고 매칭 결과만 확인합니다.
- `--report`: 매칭 결과 CSV 저장 경로입니다.

## 주의

- 번역문에 모델이 제안한 문구가 정확히 존재해야 링크가 삽입됩니다.
- 여러 번 반복되는 문구는 첫 번째로 매칭된 문단 안에서 처리됩니다.
- 문서의 복잡한 서식은 최대한 유지하지만, 링크가 들어가는 문단 내부의 일부 run은 분할될 수 있습니다.
