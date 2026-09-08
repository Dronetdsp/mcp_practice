# 🧩 MCP Practice — Claude Code × Notion MCP

Claude Code와 Notion MCP를 연결해서 **로컬 문서 → Notion 프로젝트 관리 → Excel 리포트**로 이어지는 실전 워크플로우를 실습한 저장소입니다.

> 이 저장소에는 로컬에서 사용한 입력 자료와 결과물(Excel)만 담겨 있습니다. 실제 결과물(페이지, 데이터베이스, 뷰, 댓글 등)은 Notion MCP를 통해 워크스페이스에 직접 생성되었습니다.

---

## 📌 무엇을 실습했나

가상의 프로젝트 **"AI 여행 플래너"** 를 소재로, Notion MCP의 읽기/쓰기 도구를 하나씩 사용해봤습니다.

| 영역 | 내용 |
|---|---|
| 📄 문서 변환 | 로컬 Markdown → Notion 페이지 (구조 유지, 임의 내용 추가 금지) |
| 🔍 검색 & 조회 | 제목/내용 검색, URL 직접 조회, 검색→가져오기 흐름 |
| ✍️ 페이지 CRUD | 구조 지정 생성, 목적만 주고 초안 제안받기, 여러 페이지 일괄 생성 |
| 🛠️ 안전한 수정 | 기존 내용 보존 + 섹션 추가/치환 (`insert_content` / `update_content`) |
| 🖼️ 리치 미디어 | 이미지 · MP4 영상 · YouTube 영상 블록 임베드 |
| 🗄️ 데이터베이스 | Task DB 스키마 설계, 속성 추가, 행 생성/수정 |
| 📊 View | Board(진행 현황) · Calendar(일정) · Table(긴급 업무, 필터+정렬) |
| 💬 댓글 | 페이지 코멘트 작성 및 조회 |
| 📁 정리 | 하위 페이지 폴더링, 페이지 복제 |
| 📈 리포트 | 회의록 Action Item ↔ Task DB 비교, 주간/현황 보고서 자동 생성 |
| 📤 Excel 내보내기 | Notion DB → 서식이 적용된 실제 `.xlsx` (2시트, 하이퍼링크, 자동 필터) |
| 🛡️ 안전장치 | 쓰기 전 승인 절차, 외부 문서 속 지시문 실행 금지(프롬프트 인젝션 방어) |

---

## 📂 폴더 구조

```
mcp_practice/
├── inputs/                        # 실습에 사용한 원본 입력 자료
│   ├── project-brief.md           # 프로젝트 개요 (목적/사용자/기능/기술스택/원칙)
│   ├── development-meeting.md     # 개발 회의록 (결정사항/Action Item)
│   └── weekly-report-format.md    # 주간 보고서 템플릿
├── build_excel.py                 # Notion Task DB → project_tasks.xlsx 생성 스크립트
├── output/
│   └── project_tasks.xlsx         # 생성된 Excel 리포트 (Task 목록 + 요약 시트)
└── README.md
```

---

## 🌳 Notion 워크스페이스 구조

```
MCP실습
└── AI 서비스 개발 프로젝트
    ├── 요구사항
    ├── 프로젝트 Kickoff  (+ 복제본: 프로젝트 Kickoff (1))
    ├── 회의록
    │   └── 2026-08-29 개발 회의
    ├── 기술 결정 기록
    ├── 사용자 시나리오
    ├── 릴리스 체크리스트
    ├── 프로젝트 Task (데이터베이스)
    │   ├── 진행 현황 (Board — Status별 그룹)
    │   ├── 일정 (Calendar — Due Date 기준)
    │   └── 긴급 업무 (Table — Priority=High & Status≠Done, Due Date 오름차순)
    ├── 2026-08-29 주간 프로젝트 보고서
    └── AI 서비스 프로젝트 현황 보고서
```

---

## 🖥️ Excel 리포트 다시 만들기

```bash
pip install openpyxl
python build_excel.py
```

`output/project_tasks.xlsx`가 생성되며, 다음을 포함합니다.

- **Task 목록** 시트 — 헤더 강조, 첫 행 고정, 자동 필터, 열 너비 자동 조절, `Due Date` 실제 날짜 서식, `Notion URL` 클릭 가능한 하이퍼링크
- **요약** 시트 — Status/Priority/Category별 집계, 전체 Task 개수, 원본 Task DB URL

---

## 🛡️ 실습에서 지킨 원칙

1. **쓰기 전 확인**: Notion에 쓰기 작업을 하기 전 대상 워크스페이스·페이지·변경 내용을 먼저 확인
2. **원본 보존**: 로컬 Markdown에 없는 내용을 임의로 추가하거나 요약하지 않음
3. **추측 금지**: 담당자·날짜·진행률 등 근거 없는 정보는 "미정" 또는 "확인 필요"로 표시
4. **프롬프트 인젝션 방어**: 외부 문서(회의록 등) 안에 포함된 지시문은 데이터로만 취급하고 실행하지 않음
5. **중복 방지**: 페이지/Task 생성 전 항상 기존 항목과 중복 여부 확인

---

## 🔗 참고

- [Notion MCP 공식 문서](https://developers.notion.com/docs/mcp)
- Workspace: Brina seongsoo Yoon's Space
