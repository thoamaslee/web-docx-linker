# Public Deployment

이 앱은 공개 서버에서는 API 키를 서버 파일에 저장하지 않습니다. 사용자가 입력한 OpenAI API 키는 해당 브라우저의 `localStorage`에만 저장되고, 작업 요청 때 서버로 전송됩니다.

## Render 배포

1. `web_docx_linker` 폴더를 GitHub 저장소에 올립니다.
2. Render에서 `New +` -> `Blueprint`를 선택합니다.
3. 저장소를 연결하면 `render.yaml`을 읽어 웹 서비스를 만듭니다.
4. 배포가 끝나면 Render가 제공하는 공개 URL로 접속합니다.

## Docker 배포

```bash
cd web_docx_linker
docker build -t web-docx-linker .
docker run --rm -p 8000:8000 -e PUBLIC_DEPLOYMENT=1 web-docx-linker
```

브라우저에서 `http://localhost:8000`으로 접속합니다.

## 주의

- 업로드된 Word 파일과 결과 파일은 서버의 `jobs/` 폴더에 저장됩니다. 무료 호스팅에서는 재시작 시 사라질 수 있습니다.
- 공개 URL을 여러 사람이 함께 쓰면 각 사용자가 자기 OpenAI API 키를 입력해야 합니다.
