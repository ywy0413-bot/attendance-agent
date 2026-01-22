# 근태/휴가 이메일 분석 Agent

Microsoft 365 Outlook의 근태/휴가 이메일을 자동으로 분석하여 매일 아침 8시에 엑셀 보고서를 발송하는 Agent입니다.

## 기능

- **이메일 자동 수집**: Outlook 특정 폴더에서 `[휴가신고]`, `[근태공유]` 이메일 수집
- **자동 분류**: 제목과 본문을 분석하여 4가지 카테고리로 분류
  - 휴가신고
  - 출근지연
  - 외출
  - 조기퇴근
- **정보 추출**: 신청자, 날짜, 시간, 사유 자동 추출
- **엑셀 보고서**: 시트별로 정리된 일일 보고서 생성
- **자동 발송**: 매일 오전 8시 지정된 수신자에게 이메일 발송

## 프로젝트 구조

```
attendance-agent/
├── src/
│   ├── auth/
│   │   └── graph_auth.py          # Microsoft Graph 인증
│   ├── email/
│   │   ├── email_client.py        # 이메일 읽기/전송
│   │   ├── email_parser.py        # 이메일 파싱
│   │   └── email_classifier.py    # 이메일 분류
│   ├── models/
│   │   ├── attendance.py          # 근태 데이터 모델
│   │   └── vacation.py            # 휴가 데이터 모델
│   ├── report/
│   │   └── excel_generator.py     # 엑셀 보고서 생성
│   ├── config.py                  # 환경 설정
│   └── main.py                    # 메인 처리 로직
├── tests/                         # 테스트 코드
├── scripts/
│   └── local_test.py              # 로컬 테스트 스크립트
├── docs/
│   └── IT부서_Azure_AD_앱등록_요청서.md
├── function_app.py                # Azure Functions 진입점
├── host.json                      # Azure Functions 설정
├── requirements.txt
└── README.md
```

## 설치

### 1. 의존성 설치

```bash
pip install -r requirements.txt
```

### 2. 환경 변수 설정

`.env.example`을 복사하여 `.env` 파일을 생성하고 값을 입력합니다:

```bash
cp .env.example .env
```

```env
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-client-id
AZURE_CLIENT_SECRET=your-client-secret
TARGET_MAILBOX=attendance@company.com
TARGET_FOLDER=근태휴가
REPORT_RECIPIENTS=your-email@company.com
```

### 3. Azure AD 앱 등록 (IT 부서 요청)

`docs/IT부서_Azure_AD_앱등록_요청서.md`를 참고하여 IT 부서에 앱 등록을 요청합니다.

필요한 권한:
- `Mail.Read` (Application)
- `Mail.Send` (Application)
- `User.Read.All` (Application)

## 로컬 테스트

### 기능 테스트 (Azure 연결 없이)

```bash
python scripts/local_test.py
```

### 단위 테스트

```bash
pytest tests/ -v
```

### Azure Functions 로컬 실행

```bash
# Azure Functions Core Tools 필요
func start
```

## Azure Functions 배포

### 1. Azure Functions 리소스 생성

```bash
# Azure CLI 로그인
az login

# 리소스 그룹 생성
az group create --name rg-attendance-agent --location koreacentral

# Storage Account 생성
az storage account create --name stattendanceagent --resource-group rg-attendance-agent --location koreacentral --sku Standard_LRS

# Function App 생성
az functionapp create --name func-attendance-agent --resource-group rg-attendance-agent --storage-account stattendanceagent --consumption-plan-location koreacentral --runtime python --runtime-version 3.11 --functions-version 4
```

### 2. 환경 변수 설정

```bash
az functionapp config appsettings set --name func-attendance-agent --resource-group rg-attendance-agent --settings \
    AZURE_TENANT_ID=xxx \
    AZURE_CLIENT_ID=xxx \
    AZURE_CLIENT_SECRET=xxx \
    TARGET_MAILBOX=xxx \
    REPORT_RECIPIENTS=xxx
```

### 3. 배포

```bash
func azure functionapp publish func-attendance-agent
```

## API 엔드포인트

| 엔드포인트 | 설명 |
|------------|------|
| Timer Trigger | 매일 08:00 자동 실행 |
| `GET /api/run` | 수동 실행 |
| `GET /api/health` | 상태 확인 |

## 엑셀 보고서 형식

### 시트 구성

1. **요약**: 당일 처리 건수
2. **휴가신고**: No, 신청자, 부서, 휴가일자, 휴가종류, 사유, 이메일수신시간
3. **근태공유_출근지연**: No, 신청자, 부서, 날짜, 시작시간, 종료시간, 사유, 이메일수신시간
4. **근태공유_외출**: 동일 구조
5. **근태공유_조기퇴근**: 동일 구조

## 이메일 형식 요구사항

### 휴가신고
- 제목: `[휴가신고]` 포함
- 본문: 신청자, 날짜, 휴가종류, 사유

### 근태공유
- 제목: `[근태공유]` 포함
- 본문:
  - "출근지연", "지각" 등 → 출근지연으로 분류
  - "외출", "외근" 등 → 외출로 분류
  - "조기퇴근", "조퇴" 등 → 조기퇴근으로 분류

## 라이선스

Private - 내부 사용 전용
