# 개인정보 가리개 — iPhone Duo 스타일 DESIGN.md

> 접으면 아이폰, 펼치면 두 화면.
> 애플 첫 폴더블 "아이폰 듀오"의 제품 페이지 감성(하얀 갤러리, 조용한 검정 글씨, 파란 알약 버튼, 28px 둥근 카드)을
> 펼치면 두 화면이 나란히 뜨는 폴더블 구조 위에 옮긴 데스크톱 도구 화면.

**테마:** 스타 화이트(라이트) 기본 + 나이트 스카이(다크) 지원
**구현:** pywebview + HTML/CSS/JS, 인터넷 없이 동작 (CDN·외부 주소 금지)
**기준 자료:** 아이폰 듀오 제품 페이지 스타일 레퍼런스(Gallery White / Studio Mist / Ink / Pricing Blue, 28px 카드, 그림자 없음). 제품 광고 페이지 크기(80px 제목 등)는 도구 화면에 맞게 줄였다.

---

## 1. 핵심 컨셉: 펼친 화면과 접힌 화면

### 펼친 상태 (창 너비 900px 이상) — 아이폰 듀오 내부 화면
- 창이 **두 패널로 나뉜다**: 왼쪽 **가리기**, 오른쪽 **되돌리기**.
- 아이폰 듀오의 멀티태스킹처럼 두 작업을 동시에 보며 쓴다.
  가리기 결과를 복사 → AI → 답변을 오른쪽에 바로 붙여 복원.
- 두 패널 사이에는 **힌지(접히는 선)**: 24px 간격 가운데에 1px `--hairline` 세로선 하나. 장식 없음.
- 기본 창 크기 **1180 × 860** (아이폰 듀오 펼친 화면처럼 거의 정사각에 가까운 가로형).

### 접힌 상태 (창 너비 900px 미만) — 아이폰 듀오 외부 화면
- 한 패널만 보이고, 위쪽 세그먼트 컨트롤 [가리기 | 되돌리기]로 전환한다.
- 최소 창 크기 420 × 700.
- 펼침↔접힘 전환 시 패널이 0.4초 동안 부드럽게 재배치 (`cubic-bezier(0.32, 0.72, 0, 1)`).

---

## 2. 색

### 스타 화이트 (라이트, 기본)

| 이름 | 값 | 토큰 | 용도 |
|---|---|---|---|
| Gallery White | `#ffffff` | `--white` | 카드 바탕, 상단 바 |
| Studio Mist | `#f5f5f7` | `--mist` | 패널 바탕 (카드가 놓이는 회색 띠) |
| Paper Frost | `#fafafc` | `--frost` | 흐림 처리된 상단 바, 시트 바탕 |
| Hairline Silver | `#d6d6d6` | `--hairline` | 1px 구분선, 힌지 선 |
| Control Gray | `#e6e6e8` | `--control` | 세그먼트 바탕, 비활성 버튼 |
| Ink | `#1d1d1f` | `--ink` | 제목, 본문 |
| Slate | `#707070` | `--slate` | 보조 설명, 작은 안내 글 |
| Steel | `#86868b` | `--steel` | 입력칸·보조 버튼 테두리, 비활성 표시 |
| Apple Blue | `#0066cc` | `--link` | 글자 링크, 코드 강조 글자 |
| Pricing Blue | `#0071e3` | `--action` | 주 버튼(알약) 바탕 |
| Launch Orange | `#b64400` | `--orange` | "추천", "확인 필요" 같은 작은 글자 표시 |
| Safe Green | `#248a3d` | `--safe` | "남은 개인정보 없음" 글자·점 |
| Alert Red | `#d70015` | `--danger` | "남은 개인정보 발견", 오류 |

### 나이트 스카이 (다크)

| 토큰 | 값 |
|---|---|
| `--white` | `#1d1d1f` |
| `--mist` | `#000000` |
| `--frost` | `rgba(22,22,23,0.8)` |
| `--hairline` | `#424245` |
| `--control` | `#2c2c2e` |
| `--ink` | `#f5f5f7` |
| `--slate` | `#a1a1a6` |
| `--steel` | `#6e6e73` |
| `--link` | `#2997ff` |
| `--action` | `#0071e3` |
| `--orange` | `#ff8a3d` |
| `--safe` | `#30d158` |
| `--danger` | `#ff453a` |

**색 규칙:** 화면은 거의 흑백. 파랑은 **주 버튼과 링크에만**, 초록·주황·빨강은 **상태 글자와 작은 점에만** 쓴다.

---

## 3. 글꼴

- **본문·제목:** `-apple-system, "SF Pro Display", "SF Pro Text", "Pretendard", "Apple SD Gothic Neo", "Malgun Gothic", sans-serif`
  - 맥은 시스템 글꼴이 자동 적용. 윈도우용으로 **Pretendard를 `ui/fonts/`에 넣어 `@font-face`로** 불러온다.
- **코드·파일 이름:** `ui-monospace, "SF Mono", Menlo, Consolas, monospace`

| 역할 | 크기 | 굵기 | 줄 높이 | 자간 | 토큰 |
|---|---|---|---|---|---|
| App Title (상단 바 제목) | 19px | 600 | 1.21 | 0 | `--t-bar` |
| Panel Kicker (패널 위 작은 제목) | 17px | 600 | 1.2 | 0 | `--t-kicker` |
| Panel Headline (패널 큰 제목) | 40px | 600 | 1.1 | -0.5px | `--t-headline` |
| Card Heading | 24px | 600 | 1.17 | 0 | `--t-card` |
| Body | 17px | 400 | 1.47 | -0.2px | `--t-body` |
| Body Small | 14px | 400 | 1.43 | -0.1px | `--t-small` |
| Control | 14px | 500 | 1.29 | 0 | `--t-control` |
| Label (추천, 상태) | 12px | 600 | 1.33 | 0 | `--t-label` |

- 제목은 700(Bold) 쓰지 않는다. 600까지만.
- 한글 본문은 자간을 과하게 줄이지 않는다.

---

## 4. 간격과 모양

- **기본 단위:** 4px
- **패널 안쪽 여백:** 40px (접힌 상태 20px)
- **카드 사이:** 20px
- **섹션 사이:** 48px
- **카드 안쪽 여백:** 28px (접힌 상태 20px)
- **힌지 간격:** 24px

| 요소 | 둥근 모서리 |
|---|---|
| 카드 | 28px |
| 글상자(여러 줄 입력) | 18px |
| 한 줄 입력칸 | 980px (알약) |
| 버튼 | 9999px (알약) |
| 세그먼트 바탕 / 선택칸 | 9999px |
| 상단 바(떠 있는 형태) | 20px |
| 코드 강조 칩 | 6px |
| 시트 | 28px |

### 깊이 표현
- **그림자 금지.** 흰 카드 vs 회색 띠, 1px 선으로만 층을 나눈다.
- 허용되는 유일한 테두리 효과: `0 0 0 1px` 링
  - 기본: `0 0 0 1px #e6e6e8` (`--ring`)
  - 선택·포커스: `0 0 0 1px #86868b` (`--ring-strong`)

---

## 5. 구성 요소

### 상단 바
- 높이 52px, `--frost` 바탕 + `backdrop-filter: saturate(180%) blur(20px)`, 아래 1px `--hairline`.
- 왼쪽: "개인정보 가리개" 19px/600.
- 오른쪽: 학생사전 상태 글자 14px `--slate` ("학생 30명 · 10월 18일") + **외곽선 알약 버튼** "사전 새로 만들기".
- 사전이 없으면 상태 글자가 `--danger` "학생사전 없음".
- 접힌 상태에서는 상단 바 아래에 [가리기 | 되돌리기] 세그먼트가 붙는다.

### 패널
- 바탕 `--mist`, 위에서부터: Kicker(작은 제목) → Headline(큰 제목) → 카드들.
- 왼쪽 패널: Kicker "STEP 1", Headline "가리기"
- 오른쪽 패널: Kicker "STEP 2", Headline "되돌리기"
- 각 패널은 독립적으로 세로 스크롤.

### 카드 (Feature Card)
- `--white` 바탕, 28px 모서리, 그림자 없음, 테두리 없음.
- 카드 맨 위 24px/600 제목, 아래 14px `--slate` 설명.

### 주 버튼 (Pricing Blue Pill)
- `--action` 바탕, 흰 글자 17px/500, 높이 44px, 좌우 여백 22px, 9999px 모서리.
- 누르면 투명도 0.8. 비활성은 `--control` 바탕 + `--steel` 글자.
- **한 패널에 주 버튼은 하나만** ("가명처리 실행", "복원").

### 보조 버튼 (Outlined Pill)
- 투명 바탕, `--ink` 글자 14px/500, 1px `--steel` 테두리, 높이 36px, 9999px 모서리.
- "결과 복사", "폴더 열기", "클립보드에서 바로 복원", "파일 복원", "사전 새로 만들기".

### 글자 링크
- `--link` 14px, 밑줄 없음, 끝에 " ›". 예: "변경 기록 보기 ›".

### 세그먼트 컨트롤
- 바탕 `--control`, 높이 36px, 9999px 모서리, 안쪽 3px.
- 선택칸 `--white` + `--ring` 링 (그림자 대신), 글자 14px/500.
- 방식 선택: [ AI용 md | 엑셀 원본 형식 ] — "AI용 md" 옆에 `--orange` 12px/600 글자 "추천" (배경 없음).

### 파일 선택 카드
- 카드 안 가운데 정렬: 얇은 선 문서 아이콘(인라인 SVG, `--steel`) + "파일을 끌어다 놓거나" + 보조 버튼 "파일 선택".
- 끌어다 놓는 중: 카드 테두리 2px 점선 `--action`, 바탕 `rgba(0,113,227,0.06)`.
- 선택 후: 파일 이름(모노 글꼴 14px) + 형식 표시 12px `--slate`.

### 결과 카드
- 상태 줄: 8px 동그란 점 + 17px/600 글자. 배경 상자·색 알약 쓰지 않는다.
  - 안전: `--safe` 점 + "남은 개인정보 없음"
  - 확인 필요: `--orange` 점 + "확인 필요 3건" + 링크 "목록 보기 ›"
  - 위험: `--danger` 점 + "남은 개인정보 발견" → 이때 "결과 복사" 버튼 비활성
- 미리보기 글상자: `--mist` 바탕, 18px 모서리, 안쪽 16px 20px, 15px/1.6, 높이 260px 내부 스크롤.
- 코드 칩 (`학생003`, `S003`, `[학생전화003]` 등): 글자 `--link` 600, 바탕 `rgba(0,102,204,0.08)`, 6px 모서리, 좌우 3px.
- 아래: 보조 버튼 "결과 복사", "폴더 열기".

### 글상자 (되돌리기)
- `--white` 카드 안 `--mist` 글상자, 18px 모서리, 포커스 시 `--ring-strong`.
- placeholder: "AI 답변을 붙여 넣으세요" `--steel`.

### 힌지
- 펼친 상태에서만 보임. 패널 사이 24px 간격, 가운데 1px `--hairline` 세로선, 위아래 끝은 40px씩 비움.

### 시트 (확인 필요 목록, 사전 만들기 결과)
- 화면 아래에서 올라옴, `--frost` + 흐림, 위쪽 28px 모서리, 맨 위 손잡이(36×5px).
- 펼친 상태에서는 해당 패널 너비 안에서만 올라온다 (아이폰 듀오의 한쪽 화면 시트처럼).

### 알림창
- 가운데 300px, `--white`, 28px 모서리, 링 `--ring`.
- 제목 17px/600, 내용 14px `--slate`, 버튼 알약 2개 (취소: 외곽선 / 확인: 주 버튼, 위험 동작은 `--danger` 글자 외곽선).

### 토스트
- 상단 가운데, `--ink` 바탕 + `--white` 글자 14px/500, 9999px 모서리, 1.5초 후 사라짐. 예: "복사했어요".

---

## 6. 화면 구성

### 왼쪽 패널 — 가리기
1. Kicker "STEP 1" / Headline "가리기"
2. 카드 "파일": 파일 선택 영역
3. 카드 "방식": 세그먼트 + 설명 "AI에 올릴 때는 md를 추천해요."
4. 주 버튼 "가명처리 실행" (카드 밖, 왼쪽 정렬)
5. 카드 "결과": 상태 줄 → 미리보기 → 보조 버튼들 → 링크 "변경 기록 보기 ›"

### 오른쪽 패널 — 되돌리기
1. Kicker "STEP 2" / Headline "되돌리기"
2. 카드 "AI 답변": 글상자 + 보조 버튼 "클립보드에서 바로 복원"
3. 주 버튼 "복원"
4. 카드 "복원 결과": 글상자 + 보조 버튼 "복사"
5. 카드 "확인 필요": 복원 못 한 코드 목록 (없으면 숨김)
6. 링크 "파일 복원 ›"

---

## 7. 할 것 / 하지 말 것

### 할 것
- 패널 바탕은 `--mist`, 내용은 28px 흰 카드에 담는다.
- 주 버튼은 파란 알약 하나, 나머지는 외곽선 알약.
- 상태는 작은 점 + 글자로만 표시한다.
- 펼친 상태와 접힌 상태가 모두 자연스럽게 보이도록 900px 기준으로 반응형 처리.
- 나이트 스카이(다크) 모드도 토큰만 바꿔서 지원.

### 하지 말 것
- 그라데이션, 그림자, 굵은 테두리 금지.
- 파랑을 큰 바탕색으로 쓰지 않는다.
- 상태를 색 배경 상자나 색 알약으로 감싸지 않는다.
- 28px 카드 모서리를 8~12px로 줄이지 않는다.
- 제목 Bold(700) 금지, 600까지만.
- 애플 로고, 아이폰 듀오 제품 사진, 외부 아이콘 폰트, CDN 금지.

---

## 8. CSS 토큰 (시작점)

```css
@font-face {
  font-family: "Pretendard";
  src: url("fonts/PretendardVariable.woff2") format("woff2");
  font-weight: 100 900;
}

:root {
  --white: #ffffff;
  --mist: #f5f5f7;
  --frost: rgba(250,250,252,0.85);
  --hairline: #d6d6d6;
  --control: #e6e6e8;
  --ink: #1d1d1f;
  --slate: #707070;
  --steel: #86868b;
  --link: #0066cc;
  --action: #0071e3;
  --orange: #b64400;
  --safe: #248a3d;
  --danger: #d70015;
  --code-tint: rgba(0,102,204,0.08);

  --ring: 0 0 0 1px #e6e6e8;
  --ring-strong: 0 0 0 1px #86868b;

  --font: -apple-system, "SF Pro Display", "SF Pro Text", "Pretendard", "Apple SD Gothic Neo", "Malgun Gothic", sans-serif;
  --mono: ui-monospace, "SF Mono", Menlo, Consolas, monospace;

  --r-card: 28px;
  --r-textarea: 18px;
  --r-pill: 9999px;
  --r-bar: 20px;
  --r-chip: 6px;

  --pad-panel: 40px;
  --pad-card: 28px;
  --gap: 20px;
  --gap-section: 48px;
  --hinge: 24px;
  --fold-point: 900px;
  --ease-fold: cubic-bezier(0.32, 0.72, 0, 1);
}

@media (prefers-color-scheme: dark) {
  :root {
    --white: #1d1d1f;
    --mist: #000000;
    --frost: rgba(22,22,23,0.8);
    --hairline: #424245;
    --control: #2c2c2e;
    --ink: #f5f5f7;
    --slate: #a1a1a6;
    --steel: #6e6e73;
    --link: #2997ff;
    --orange: #ff8a3d;
    --safe: #30d158;
    --danger: #ff453a;
    --code-tint: rgba(41,151,255,0.16);
    --ring: 0 0 0 1px #424245;
    --ring-strong: 0 0 0 1px #6e6e73;
  }
}

body {
  margin: 0;
  background: var(--mist);
  color: var(--ink);
  font: 400 17px/1.47 var(--font);
  -webkit-font-smoothing: antialiased;
}

/* 펼친 상태: 두 패널 + 힌지 */
.panels { display: grid; grid-template-columns: 1fr var(--hinge) 1fr; }
.hinge { position: relative; }
.hinge::after {
  content: ""; position: absolute; left: 50%; top: 40px; bottom: 40px;
  width: 1px; background: var(--hairline);
}

/* 접힌 상태: 한 패널 + 세그먼트 전환 */
@media (max-width: 899px) {
  .panels { grid-template-columns: 1fr; }
  .hinge { display: none; }
}
```
