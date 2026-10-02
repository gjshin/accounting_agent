# Word 양식

검토 의견서와 메모를 Word로 바꿀 때(`tools/md_to_docx.py`) 쓰는 양식 폴더다.

- 회사(회계법인) 양식을 쓰려면 이 폴더에 `reference.docx`로 저장한다. 변환할 때 그 파일의 글꼴, 제목 스타일, 머리글·바닥글, 여백을 그대로 쓴다.
- `reference.docx`가 없으면 맑은 고딕 기본 양식(머리글 "대외비 | 회계위키·법령 근거 검토", 쪽 번호)을 자동으로 만든다.
- 양식 파일의 스타일 이름은 Word 기본 이름(Title, Heading 1~3, Block Text, Footnote Text 등)을 유지해야 적용된다. 바탕 문서는 `python3 -c "import pypandoc,subprocess;subprocess.run([pypandoc.get_pandoc_path(),'-o','templates/reference.docx','--print-default-data-file','reference.docx'])"`로 받아 Word에서 스타일만 고쳐 저장하면 된다.
