# Word AI Assistant

MS Word에서 사이드패널로 동작하는 AI 문서 작성 어시스턴트입니다.
템플릿 Word 문서를 열고 raw 데이터를 채팅으로 전달하면, AI가 문서 구조를 이해하고 알아서 채워넣습니다.

## 동작 원리

```
┌──────────────────────────────────────┐
│  MS Word                             │
│  ┌──────────┐  ┌──────────────────┐  │
│  │ 문서영역  │  │ 사이드패널(웹뷰) │  │
│  │ .docx    │◄►│ React 웹앱      │  │
│  │          │  │ + Office.js     │  │
│  └──────────┘  └───────┬──────────┘  │
│                        │ fetch()     │
└────────────────────────┼─────────────┘
                         ▼
                  AI Server (on-prem)
                  OpenAI 호환 API
```

- **사이드패널 = 웹앱**: Word가 내장 브라우저로 우리가 만든 웹페이지를 열어줌
- **Office.js**: MS 제공 JavaScript API로 문서 내용을 읽고 쓸 수 있음
- **AI 직접 호출**: 사이드패널에서 AI 서버로 직접 요청 (중간 백엔드 없음)
- **서버는 정적 파일만 서빙**: 로직은 전부 브라우저(Word 내부)에서 실행

## 주요 기능

- **에이전트형 AI**: 문서 구조(단락, 표, 스타일)를 깊이 분석하고, 대화하면서 판단해서 수정
- **스트리밍 응답**: AI 응답이 실시간으로 토큰 단위로 표시
- **멀티턴 대화**: 이전 대화 맥락을 유지하면서 점진적으로 작업
- **슬래시 명령어**: `/clear`, `/summary`, `/help`
- **설정 가능한 AI 백엔드**: OpenAI 호환 API URL, API Key, 모델명 자유 설정

## 빠른 시작 (개발)

```bash
# 의존성 설치
npm install

# 개발 서버 시작 (https://localhost:3000)
npm run dev
```

### Mac Word에서 테스트

```bash
# manifest를 Word 사이드로드 폴더에 복사
mkdir -p ~/Library/Containers/com.microsoft.Word/Data/Documents/wef
cp manifest.xml ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/

# Word 재시작 → 삽입 → 추가 기능 → AI Assistant
```

### Windows Word에서 테스트

1. Word 열기 → **삽입** → **추가 기능** → **내 추가 기능** → **업로드**
2. `manifest.xml` 선택

## 배포 (Docker)

```bash
# 1. SSL 인증서 준비
#    회사 인증서가 있는 경우:
cp /path/to/cert.pem certs/cert.pem
cp /path/to/key.pem certs/key.pem

#    없는 경우 (자체 서명 인증서 생성):
./init-certs.sh

# 2. manifest.xml의 URL을 실제 서버 주소로 변경
#    SourceLocation, IconUrl 등의 localhost:3000을 실제 주소로 교체

# 3. 서버 시작
docker compose up -d --build
```

서버가 `https://your-server/taskpane.html`에서 서빙됩니다.

## 다른 PC에 설치

### 방법 1: 수동 설치 (개인별)
1. `manifest.xml` 파일 전달
2. Word → **삽입** → **추가 기능** → **내 추가 기능** → **업로드**

### 방법 2: 네트워크 공유 (팀 단위)
1. `manifest.xml`을 네트워크 공유 폴더에 배치
2. 각 PC에서 Word → **파일** → **옵션** → **보안 센터** → **신뢰할 수 있는 추가 기능 카탈로그**에 공유 폴더 경로 추가

### 방법 3: Microsoft 365 관리센터 (조직 전체)
1. 관리센터 → **설정** → **통합 앱** → **사용자 지정 앱 업로드**
2. `manifest.xml` 업로드 → 배포 대상 지정
3. 모든 직원의 Word에 자동으로 표시됨

## 설정

사이드패널의 **Settings** 탭에서 설정:

| 항목 | 설명 | 예시 |
|------|------|------|
| AI Server URL | OpenAI 호환 API 엔드포인트 | `https://openrouter.ai/api/v1/chat/completions` |
| API Key | 인증 키 (선택) | `sk-...` |
| Model Name | 사용할 모델 ID | `z-ai/glm-5`, `openai/gpt-4o` |

## 슬래시 명령어

입력창에서 `/`를 입력하면 명령어 팔레트가 표시됩니다.

| 명령어 | 설명 |
|--------|------|
| `/clear` | 대화 내역 전체 초기화 |
| `/summary` | 현재 문서 내용 요약 |
| `/help` | 명령어 목록 표시 |

## 사용 예시

1. 표가 있는 분기 보고서 템플릿을 Word에서 열기
2. 사이드패널에서 AI Assistant 열기
3. 채팅에 raw 데이터 입력:
   ```
   매출: Q3 25억, Q4 31억
   영업이익: Q3 5억, Q4 7억
   순이익: Q3 3.5억, Q4 5.2억
   ```
4. AI가 문서 구조를 분석하고, 적절한 셀에 데이터를 채워넣음

## 기술 스택

- **Frontend**: React 19 + TypeScript
- **Office API**: Office.js (Word API)
- **Bundler**: Webpack 5
- **서버**: nginx (Docker)
- **AI Protocol**: OpenAI Chat Completions API (streaming)

## 프로젝트 구조

```
├── manifest.xml                # Office Add-in 매니페스트
├── Dockerfile                  # Docker 빌드
├── docker-compose.yml          # Docker 실행
├── nginx.conf                  # nginx HTTPS 설정
├── init-certs.sh               # 자체 서명 인증서 생성
├── src/taskpane/
│   ├── index.html              # HTML 진입점
│   ├── index.tsx               # React 마운트
│   ├── App.tsx                 # 메인 앱 (시스템 프롬프트, 멀티턴)
│   ├── components/
│   │   ├── ChatPanel.tsx       # 채팅 UI (스트리밍, 슬래시 명령)
│   │   ├── MessageBubble.tsx   # 메시지 표시
│   │   └── SettingsPanel.tsx   # 설정 화면
│   ├── services/
│   │   ├── aiClient.ts         # AI API 호출 (SSE 스트리밍)
│   │   ├── wordDocument.ts     # Office.js 문서 읽기/쓰기
│   │   └── settings.ts         # 설정 저장
│   └── styles/
│       └── app.css             # 스타일
└── assets/                     # 아이콘
```
