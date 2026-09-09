# Manager Format 구현·배포 검증 보고서

2026-09-09. 기능 커밋: `5e5f33f` (main → origin/main). Cloudflare Pages 체크 성공.
실제 확인 주소: https://cbd-generator.pages.dev/#comparison

## 결과와 계산 정책 예외

기존 상세 비교 화면과 두 기존 Excel 추출을 유지하고 다음 버튼을 추가했다.

- Export Current Style – Manager Format
- Export All Styles – Manager Format

첨부 13개 STYLE를 모두 분석하고 동적 행 생성기로 출력했다. 신규는 A:I, 기존은 Q:Y, 중앙 K:O는 신규 − 기존이다. Manager에는 STYLE MATCH 시트가 없고 기존 상세 출력에는 그대로 있다.

**첨부 인쇄 Total과의 완전 일치는 성립하지 않는다.** 첨부 13개 시트에는 SEWING THREAD 자재 행 Extended Cost 없이 별도 소계로만 $0.25가 입력되어 있다. 기존 상세 계산은 원본 소계를 참고용으로 제외한다. 신규 원본 Total에는 그룹 반올림 차이도 있다. 기존 계산을 바꾸지 않는 요구를 우선하여 비교 Total을 보존했고, Manager에 원본 소계·원본 Total·차이를 참고 정보로 별도 표시했다. 이 차이를 숨기거나 “원본 Total 일치 통과”로 판정하지 않았다. 계산 정책 선택은 사용자에게 질문했으며, 이후 기존 계산 보존 방향으로 계속 진행한다고 알렸다.

EVO 예: 신규 원본 인쇄 Total 5.2085 / 기존 상세와 Manager 비교 Total 4.958647363005. 기존 원본 인쇄 Total 2.773355128 / 비교 Total 2.523355128. FINAL FOB는 신규 9.39 / 기존 6.80으로 원본 및 기존 상세와 일치한다.

## 전 시트 분석

`reports/manager-template-analysis.md`에 13개 시트별 그룹 시작 행, 자재 수, 빈 상대편 수, Total/FOB 위치를 기록했다. 상세 원본 스타일·병합·열 너비·인쇄 정보는 로컬 `reports/manager-template-structure.json`에 있다.

공통: A:I/K:O/Q:Y, J/P 좁은 공백, C2:G2 및 S2:V2 모델명 병합, B4/R4 REFERENCES와 5행 헤더, OUTSHELL → TRIMS → SEWING THREAD → LABEL & PACKAGING → SPECIAL PROCESS 순서, 노란 소계와 청록 하단.

예외: EVO의 COOL MESH 4540 NYLON ↔ Q-SPAN, RHEON NUCKLE ↔ TPR KNUCKLE; 전 STYLE의 J RING 및 LABEL 공급처·시즌 표현 차이; KINETIC Sonic/sub의 서로 다른 자재·공백·금액; LITE 계열의 신규가 빈 TRIMS 행; PRO LITE DBK 2 원본 시트명 끝 닫는 괄호. 그룹 병합과 자재 수, 하단 위치는 STYLE마다 다르다. 원본에서 오른쪽만 잡힌 인쇄 영역은 복제하지 않았다.

로고는 공통 EMF 하나를 공유한다. 내부 기록은 실제 143×44 비트맵이며 벡터가 아니다. 해당 비트맵을 직접 추출해 흰 배경을 투명화한 PNG로 재사용한다. 원본에 없는 고해상도 디테일을 만들어내지 않았다. 각 시트 정확히 2개다.

## 데이터·수식·행 배치

- 두 렌더러가 같은 저장된 comparison model, 그룹 해결 함수, COST_GROUPS를 사용한다. 문자열 매칭·FOB 추론·Extended Cost 재계산을 새로 만들지 않았다.
- 수동 잠금과 자동 매칭의 확정된 cluster를 사용한다. 관계가 있는 대체 자재는 같은 행에 배치한다. 불확실하거나 연결되지 않은 자재를 임의로 채우지 않는다.
- 명시적 manualOrder, 기준 원본 행 순서, 근거 없는 신규 전용 행의 그룹 끝 배치 순서를 적용한다.
- 신규 전용은 Q:Y 실제 빈 셀, 기존 전용은 A:I 실제 빈 셀을 유지한다. 모든 원본 ID의 존재와 중복을 출력 전에 검사한다.
- N:1 및 저장된 1:N 관계는 anchor와 바로 아래 구성 행으로 펼친다. 각 원본 Extended Cost는 한 번만 출력한다. K:N은 비우고 anchor O만 SUM(신규 구성 행) − SUM(기존 구성 행)이다. 신규 1:N 연결 UI나 별도 자동 분리 매칭은 추가하지 않았다.
- 일반 행: K=E−U, L=K/U, M=F−V, N=M/V, O=H−X. IF로 공란·0 분모를 처리하고 IFERROR는 쓰지 않는다. 기존 전용 변화율은 −100%, 신규 전용 및 기존 0 분모의 비율은 공란이다. Unit이 다르면 K:N을 비운다. SPECIAL PROCESS는 비교 수식을 넣지 않는다.
- 그룹 소계·Total·FOB의 O는 신규 − 기존. 그룹과 Total 수식은 생성된 실제 행을 참조한다. 원본 Extended Cost를 보존하여 PCS·고정값·특수 계산의 기존 정책을 유지한다.
- 원본 FINAL FOB만 사용하며 ratio-derived FOB는 공란 처리한다. Labor/Overhead/Profit과 추가 요약은 파싱된 값만 표시하며, 하단 요약의 원본 I/Y Remark도 유지한다. Material cost ratio는 비교 Total / FINAL FOB다.

## GROUP 이동 검증

`referenceOriginalGroup`, `comparisonOriginalGroup`, `effectiveGroup`, `groupChangedFrom`, `groupChangedTo`, `groupChangeSource`를 보존한다. 원본 자재 group은 변경하지 않는다. 기존 상세 상태의 방향 표시를 유지한다. Manager 오른쪽 Y Remark에는 기존 Remark 뒤에 `[GROUP MOVE: OUTSHELL → TRIMS]`가 표시된다. 신규 원본 GROUP도 이동한 경우 I에도 근거를 남긴다.

단위 테스트: N:1 기존 금액 2+3=5를 OUTSHELL에서 TRIMS로 이동하고 신규 4를 같은 블록에 놓았다. OUTSHELL 기존 1, TRIMS 기존 5/신규 4, 양 시즌 전체 비교 Total 6이 유지된다. 재자동 매칭·저장 복원 후에도 이동 및 잠금이 유지된다.

실제 서비스: EVO 첫 관계를 TRIMS로 변경하고 Undo → Redo → 재로드했다. 기존 0.46972254와 신규 0.683401454505가 TRIMS 소계로 이동한다. Current/All Manager와 Current/All 상세의 모든 비용 그룹 소계·전체 비교 Total·FINAL FOB를 다운로드 파일끼리 대조해 일치했다.

## 디자인·인쇄 검증

99CCFF 헤더, C0C0C0 모델명, FFFF00 소계, CCFFFF 하단, FFCCFF 변경 강조, 신규 빨간 모델명/기존 검정 모델명, Arial 본문 8pt/주요 헤더 10pt Bold, 검정 내부선/굵은 외곽선, 지정 A:Y 열 너비, 음수 금액·비율의 빨간 괄호 서식을 적용했다. 자재·Remark는 줄바꿈하고 숫자는 오른쪽 정렬하며 줄바꿈하지 않는다.

13개 STYLE의 자연 크기 100% PDF/PNG를 모두 열어 확인했다. 별도 A4 가로 PDF는 13페이지이며 좌측·중앙·우측·로고·하단이 모두 들어간다. PDF 텍스트 좌표 검사로 페이지 밖 텍스트, 숫자와 다른 텍스트의 겹침, 오류 토큰이 없음을 확인했다. 최종 워크북 인쇄 영역은 A1:Y마지막행, 가로 1페이지/세로 자동, 화면 105%다. 검증기는 인쇄 설정을 유지한 PDF와 100% 자연 크기 PDF를 별도로 생성한다.

## 테스트·회귀·서비스 재추출

- `npm run test`: 12개 테스트 파일, **139개 테스트 통과**.
- `npm run build`: 성공. 기존 대형 번들 경고만 존재.
- 생성 Current: 1 STYLE, 135개 수식. All: 13 STYLE, 1,582개 수식.
- 신규 원본 자재 323개, 기존 351개가 각 STYLE에서 정확히 한 번씩 출력됨을 검사했다.
- 원본 자재 행 기반 비교 Total 및 최종 GROUP 소계는 기존 상세와 일치한다. 원본 인쇄 Total 예외는 위에 명시했다.
- LibreOffice 실제 재계산: Current/All 총 **1,717개 수식**, 오류 0, 저장된 결과와 독립 재계산값 전부 일치. #VALUE!, #REF!, #DIV/0!, #NAME?, #N/A, 순환 참조가 발견되지 않았다. 차이 0의 인상률 캐시도 0%로 검증했다.
- 기준 커밋 `61f9719`의 기존 렌더러와 현재 렌더러를 동일한 13 STYLE 데이터로 실행: Current/All의 worksheet model(값·수식·서식·시트명·인쇄 설정) 전체 동일.
- 실제 서비스에서 첨부 자료의 양 시즌·확정 관계를 담은 검증 상태를 새 브라우저의 저장소에 로드했다. 개인 기존 저장 데이터는 사용하지 않았다. 네 추출 버튼을 실행했고 Manager 1/13시트, 상세 2/14시트를 실제 다운로드했다. 브라우저 오류 0.
- 서비스 다운로드 파일도 독립 재계산·인쇄 검사를 수행했다. 실제 서비스 URL: https://cbd-generator.pages.dev/#comparison
- Cloudflare Pages main 체크: 성공. 기능 커밋 `5e5f33f`. GitHub Pages workflow 수정·수동 실행 없음.

## 수정 파일

| 파일 | 내용 |
|---|---|
| src/comparison/managerExporter.ts | 독립 Manager 렌더러와 다운로드 |
| src/comparison/model.ts | 두 렌더러의 공통 comparison model 및 비용 그룹 |
| src/comparison/relations.ts | 단일/분리 관계 원본 ID 접근 |
| src/comparison/groups.ts | 이동 방향·출처 메타데이터, 분리 관계 그룹 합계 |
| src/comparison/types.ts | 이동 메타데이터, sourceDate, 선택적 1:N/manualOrder |
| src/comparison/exporter.ts | 공통 모델 연결; 기존 형식·출력 유지; 저장된 분리 관계 합산 |
| src/comparison/matcher.ts | 분리 관계 잠금·무결성 보존; 기존 점수/매칭 규칙 유지 |
| src/comparison/parser.ts | 원본 Date 및 하단 요약 Remark 메타데이터 보존 |
| src/comparison/ComparisonWorkspace.tsx | 신규 Current/All 버튼, 추출 오류 표시 |
| src/comparison/managerExporter.test.ts | 공란·0·N:1·1:N·그룹·금액·서식·요약 Remark 회귀 테스트 |
| src/comparison/ComparisonWorkspace.test.tsx | 신규 메뉴의 독립 동작·필터 비의존 테스트 |
| src/comparison/assets/fly-logo.png, fly-logo.ts | 원본 내부 실제 로고 |
| scripts/*manager* | 템플릿 분석, 로고 추출, 실제 파일 생성, 브라우저·수식·인쇄·회귀 검증 |
| reports/manager-template-analysis.md | 13개 원본 시트 분석 |

## 검토 파일

- 실제 서비스 Current: `reports/manager-live-current.xlsx`
- 실제 서비스 All: `reports/manager-live-all.xlsx`
- 실제 서비스 상세 Current/All: `reports/manager-live-detail-current.xlsx`, `reports/manager-live-detail-all.xlsx`
- A4 인쇄 PDF: `reports/manager-render/manager-live-all-print.pdf`
- 100% 전체 시트 PDF: `reports/manager-render/manager-live-all-100pct.pdf`
- 서비스 다운로드 검증: `reports/manager-live-verification.json`
- 서비스 합계 대조: `reports/manager-live-financials.json`
- 독립 재계산: `reports/manager-render/recalculation.json`
- 전체 레이아웃: `reports/manager-layout-verification.json`
- 테스트/빌드 로그: `reports/manager-tests.log`, `reports/manager-build.log`
