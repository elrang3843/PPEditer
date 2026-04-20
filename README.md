# PPEditer

**Planning Proposal Editor** — PPT/PPTX 기획서 편집기

---

## 소개

PPEditer는 기획서(Planning Proposal) 작성 및 편집을 위한 오픈소스 프레젠테이션 편집기입니다.  
Windows 표준 환경(메뉴, 단축키, 테마)을 준수하며, PPT/PPTX 형식 열기·편집·저장과 PDF 내보내기를 지원합니다.

---

## 기능

| 기능 | 설명 |
|------|------|
| **파일 열기** | `.pptx`, `.ppt` 파일 열기 |
| **편집** | 텍스트 편집, 슬라이드 추가/삭제/복제/순서변경 |
| **저장** | `.pptx` 형식으로 저장 |
| **PDF 내보내기** | 전체 슬라이드를 PDF로 내보내기 |
| **실행 취소/다시 실행** | Ctrl+Z / Ctrl+Y |
| **확대/축소** | 창 맞춤, 50%~200% 줌 |

---

## 단축키

| 단축키 | 동작 |
|--------|------|
| `Ctrl+N` | 새 프레젠테이션 |
| `Ctrl+O` | 파일 열기 |
| `Ctrl+S` | 저장 |
| `Ctrl+Shift+S` | 다른 이름으로 저장 |
| `Ctrl+Shift+E` | PDF 내보내기 |
| `Ctrl+Z` | 실행 취소 |
| `Ctrl+Y` | 다시 실행 |
| `Ctrl+M` | 새 슬라이드 추가 |
| `Ctrl+D` | 슬라이드 복제 |
| `Ctrl+Shift+↑/↓` | 슬라이드 순서 이동 |
| `Ctrl+=` / `Ctrl+-` | 확대 / 축소 |
| `Ctrl+0` | 창 맞춤 |
| `Esc` | 편집 취소 / 선택 해제 |

---

## 설치 및 실행

### 요구사항

- Python 3.7+
- PyQt5 >= 5.15.0
- python-pptx >= 0.6.21
- lxml >= 4.6.0
- Pillow >= 9.0.0

### 설치

```bash
pip install -r requirements.txt
```

### 실행

```bash
python main.py
# 또는 파일을 바로 열기
python main.py 파일.pptx
```

---

## 라이선스

MIT License — Copyright © 2026 Noh JinMoon

---

## 회사 정보

**HANDTECH (핸텍) — 상상공작소**  
저작권자: 노진문 (Noh JinMoon)
