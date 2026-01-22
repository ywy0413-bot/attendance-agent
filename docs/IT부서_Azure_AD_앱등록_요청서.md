# Azure AD 앱 등록 요청서

## 1. 기본 정보

| 항목 | 내용 |
|------|------|
| **앱 이름** | 근태/휴가 이메일 분석 Agent |
| **용도** | Outlook 이메일 자동 분석 및 보고서 발송 |
| **요청자** | [이름] |
| **요청일** | [날짜] |
| **담당 부서** | [부서명] |
| **연락처** | [이메일/전화번호] |

---

## 2. 앱 유형 및 인증 방식

| 항목 | 내용 |
|------|------|
| **앱 유형** | 백그라운드 서비스 (데몬 앱) |
| **인증 방식** | Client Credentials Flow (앱 전용 인증) |
| **사용자 로그인** | 불필요 (무인 실행) |

---

## 3. 필요한 Microsoft Graph API 권한

### 3.1 Application 권한 (관리자 동의 필요)

| 권한 이름 | 설명 | 사용 목적 |
|-----------|------|-----------|
| `Mail.Read` | 모든 메일함의 메일 읽기 | 근태/휴가 이메일 수집 |
| `Mail.Send` | 사용자 대신 메일 전송 | 분석 보고서 이메일 발송 |
| `User.Read.All` | 모든 사용자 프로필 읽기 | 신청자 부서 정보 조회 |

### 3.2 권한 상세 설명

1. **Mail.Read (Application)**
   - 특정 메일함([대상 이메일 주소])에서 `[휴가신고]`, `[근태공유]` 제목의 이메일만 읽습니다
   - 다른 이메일 내용은 접근하지 않습니다

2. **Mail.Send (Application)**
   - 처리 결과 보고서를 지정된 수신자에게 발송합니다
   - 발송자: [대상 이메일 주소]
   - 수신자: [보고서 수신자 이메일]

3. **User.Read.All (Application)**
   - 이메일 발신자의 부서 정보만 조회합니다
   - 다른 사용자 정보는 사용하지 않습니다

---

## 4. 보안 강화 요청 (선택)

### 4.1 Application Access Policy 적용

특정 메일함에만 접근을 제한하기 위해 [Application Access Policy](https://learn.microsoft.com/en-us/graph/auth-limit-mailbox-access) 적용을 요청합니다.

**접근 허용 메일함:**
- [대상 이메일 주소] (이메일 읽기/발송용)

### 4.2 권장 설정

```powershell
# PowerShell 예시 (IT 관리자 실행)
New-ApplicationAccessPolicy -AppId [앱_클라이언트_ID] `
    -PolicyScopeGroupId [메일그룹_또는_사용자] `
    -AccessRight RestrictAccess `
    -Description "근태/휴가 Agent 메일 접근 제한"
```

---

## 5. 인증 정보 요청

### 5.1 Client Secret 방식 (권장: 간편)

| 항목 | 요청 값 |
|------|---------|
| **유효 기간** | 24개월 |
| **설명** | 근태/휴가 Agent 인증용 |

### 5.2 인증서 방식 (대안: 보안 강화)

인증서 기반 인증을 사용할 경우, 자체 생성 인증서 업로드 가능

---

## 6. 앱 등록 후 필요한 정보

앱 등록 완료 후 아래 정보를 전달해 주세요:

| 항목 | 예시 형식 |
|------|-----------|
| **Tenant ID** | xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx |
| **Client ID** | xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx |
| **Client Secret** | xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx |

---

## 7. 앱 동작 설명

### 7.1 실행 주기
- **실행 시간**: 매일 오전 8시 (KST)
- **실행 환경**: Azure Functions (Korea Central)

### 7.2 처리 프로세스

```
1. [대상 메일함]에서 전날~당일 이메일 조회
   └─ 필터: 제목에 "[휴가신고]" 또는 "[근태공유]" 포함

2. 이메일 분류 및 정보 추출
   ├─ [휴가신고]: 신청자, 휴가일자, 휴가종류, 사유
   └─ [근태공유]: 신청자, 날짜, 시간, 사유
       ├─ 출근지연
       ├─ 외출
       └─ 조기퇴근

3. 엑셀 보고서 생성

4. [보고서 수신자]에게 이메일 발송
   └─ 첨부: 근태휴가_보고서_YYYYMMDD.xlsx
```

### 7.3 접근하는 데이터

| 데이터 | 접근 범위 | 사용 목적 |
|--------|-----------|-----------|
| 이메일 제목 | [휴가신고], [근태공유] 포함 건만 | 분류 |
| 이메일 본문 | 상동 | 정보 추출 |
| 이메일 발신자 | 상동 | 신청자 식별 |
| 사용자 부서 | 이메일 발신자만 | 보고서 작성 |

---

## 8. 문의 사항

앱 등록 관련 문의:
- 담당자: [이름]
- 이메일: [이메일]
- 전화: [전화번호]

---

## 9. 참고 문서

- [Microsoft Graph API 권한 참조](https://learn.microsoft.com/en-us/graph/permissions-reference)
- [Azure AD 앱 등록 가이드](https://learn.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)
- [Application Access Policy](https://learn.microsoft.com/en-us/graph/auth-limit-mailbox-access)
