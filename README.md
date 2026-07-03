# 🏡 Decision (가족 중대 의사결정 플랫폼) - GitHub Pages

가족의 상속, 요양, 건물/부동산, 자금 등 주요 중대 의사결정을 안건으로 등록하고, 가족 구성원의 의견·투표·승인·보류·거절 이력을 체계적으로 기록/관리하며, Google Docs 회의록을 자동 생성하는 **GitHub Pages 웹 애플리케이션**입니다.

---

## 📌 주요 특징

1. **GitHub Pages 무료 호스팅**:
   - `index.html` 단일 파일로 구성되어 GitHub Pages를 통해 무료로 웹 앱을 서비스할 수 있습니다.
2. **Google Sheets & Docs 실시간 DB 연동 (REST API)**:
   - Apps Script URL 설정 시 Google Sheets (메인 DB) 및 Google Docs (회의록 생성)와 즉시 연동됩니다.
   - URL 미설정 시 내 컴퓨터/스마트폰 자체 저장소(LocalStorage) 모드로 동작합니다.
3. **엄격한 의사결정 승인 규칙**:
   - **승인 전 실행중/완료 차단**: 최종결정권자의 `승인` 상태를 거치지 않은 안건은 `실행중` 또는 `완료` 상태로 넘어갈 수 없습니다.
   - **보류 사유 필수**: `보류` 상태 선택 시 보류 사유와 재검토 예정일 저장이 강제됩니다.
4. **모바일 / PC 앱 설치 지원**:
   - 브라우저의 [홈 화면에 추가] 또는 [앱으로 설치] 기능을 이용해 데스크톱 및 스마트폰 바탕화면에 단독 앱 아이콘을 만들어 사용할 수 있습니다.

---

## 📂 파일 구조

```
decision/
├── index.html      # GitHub Pages 엔트리 웹 UI (CSS, JS, REST API client 포함)
├── Code.gs         # Google Apps Script 백엔드 (Google Sheets CRUD & Docs 회의록 생성 API)
└── README.md       # GitHub Pages 배포 및 구글 시트 연동 가이드
```

---

## 🚀 GitHub Pages 배포 가이드 (3분 완료)

### 1단계: GitHub 저장소 업로드
1. [GitHub](https://github.com)에 로그인 후 **New Repository**를 생성합니다. (Repository name: `decision`)
2. 본 `decision` 폴더 안의 파일들 (`index.html`, `Code.gs`, `README.md`)을 생성한 저장소에 업로드(push)합니다.

### 2단계: GitHub Pages 활성화
1. GitHub 저장소 상단 **Settings > Pages** 메뉴로 이동합니다.
2. **Build and deployment > Source**에서 `Deploy from a branch`를 선택합니다.
3. **Branch**를 `main` (또는 `master`) / `/ (root)`로 설정하고 **Save**를 클릭합니다.
4. 1~2분 후 배포된 **웹 사이트 URL** (`https://<본인-아이디>.github.io/decision/`)이 생성됩니다.

---

## 📊 Google Sheets (DB) 연동 방법 (선택)

GitHub Pages 웹 사이트에서 제출한 데이터를 나만의 **Google Sheets**에 저장하려면 아래 작업을 1회만 수행하면 됩니다.

1. [Google Sheets](https://sheets.google.com)에 접속하여 새 스프레드시트를 만듭니다. (이름: `Decision_DB`)
2. 상단 메뉴 **[확장 프로그램] > [Apps Script]** 클릭.
3. 기존 코드를 지우고 본 프로젝트의 [`Code.gs`](file:///c:/Users/ldjko/Desktop/Antigravity_Making/decision/Code.gs) 파일 내용을 복사해 붙여넣습니다.
4. 우측 상단 **[배포] > [새 배포]** 클릭:
   - **유형**: 웹 앱 (Web App)
   - **다음 사용자 권한으로 실행**: `나 (My account)`
   - **액세스 권한이 있는 사용자**: `모든 사용자 (Anyone)`
5. 배포 후 생성된 **웹 앱 URL** (`https://script.google.com/macros/s/.../exec`)을 복사합니다.
6. 내 **GitHub Pages 웹 사이트**에 접속하여 우측 상단의 **[⚙️ 구글시트 연동 설정]** 버튼을 누르고 복사한 URL을 붙여넣고 저장하면 모든 데이터가 구글 시트 및 구글 문서와 실시간 연동됩니다!

---

## 🔄 의사결정 프로세스 및 승인 규칙

1. **안건 등록 (`제안`)**: 제안자, 최종결정자, 제목, 핵심내용 필수 작성.
2. **가족 의견 수렴 & 투표 (`논의중` -> `승인대기`)**: 구성원별 찬/반/보류 투표 및 의견 등록.
3. **최종 결단 (`승인` / `거절` / `보류`)**:
   - **승인**: 최종결정권자가 승인 처리.
   - **보류**: 보류 사유 및 재검토 예정일 필수 입력.
   - **거절**: 안건 기각 및 사유 기록.
4. **실행 및 완료 (`실행중` -> `완료`)**: `승인` 미거침 시 차단.
5. **Google Docs 회의록 자동 생성**: `승인`, `보류`, `거절` 안건 상세 화면에서 **[📝 회의록 자동 생성]** 클릭 시 Google Docs 회의록 발행.
