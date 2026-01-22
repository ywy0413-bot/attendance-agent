import re
import logging
from datetime import datetime, timedelta
from typing import Dict, List, Any
from collections import defaultdict

from config import Settings
from auth.graph_auth import GraphAuthenticator
from mail.email_client import EmailClient, EmailMessage
from mail.email_classifier import EmailClassifier, EmailCategory, AttendanceSubType
from mail.email_parser import EmailParser
from models.attendance import AttendanceRecord
from models.vacation import VacationRecord
from report.excel_generator import ExcelReportGenerator
from utils.employee_mapper import EmployeeMapper

logger = logging.getLogger(__name__)

# 테스트 모드 - 휴가차감 메일을 실제 직원 대신 이 주소로 발송
TEST_DEDUCTION_EMAIL = "wyyu@envision.co.kr"

# 휴가차감 메일 참조자
DEDUCTION_CC_EMAIL = "brkwon@envision.co.kr"

# 보고서 메일 참조자
REPORT_CC_EMAIL = "brkwon@envision.co.kr"


async def process_attendance_emails(settings: Settings = None) -> Dict[str, Any]:
    """
    근태/휴가 이메일 처리 메인 함수

    Args:
        settings: 애플리케이션 설정 (None이면 환경변수에서 로드)

    Returns:
        Dict: 처리 결과 통계
    """
    if settings is None:
        settings = Settings.from_env()

    settings.validate()

    logger.info("근태/휴가 이메일 처리를 시작합니다")

    # 1. Graph API 인증
    auth = GraphAuthenticator(
        tenant_id=settings.azure_tenant_id,
        client_id=settings.azure_client_id,
        client_secret=settings.azure_client_secret
    )
    graph_client = auth.get_client()

    # 2. 이메일 클라이언트 초기화
    email_client = EmailClient(graph_client, settings.target_mailbox)
    classifier = EmailClassifier()
    parser = EmailParser()

    # 3. 폴더 전체 이메일 조회 (날짜 필터 없음)
    emails = await email_client.fetch_emails(
        since=None,  # 전체 폴더 조회
        folder_name=settings.target_folder if settings.target_folder != "Inbox" else None,
        subject_filter="[휴가신고] OR [근태공유]",
        max_count=500  # 더 많은 이메일 조회
    )

    logger.info(f"{len(emails)}개의 이메일을 조회했습니다")

    # 4. 분류 및 파싱
    results = {
        'vacations': [],
        'late_arrivals': [],
        'outings': [],
        'early_leaves': [],
        'unclassified': [],
        'report_date': datetime.now()
    }

    for email in emails:
        await _process_single_email(
            email=email,
            classifier=classifier,
            parser=parser,
            email_client=email_client,
            results=results
        )

    # 5. 로그 출력
    _log_results(results)

    # 6. 이전 휴가차감 이력 조회 (근태 자동화 폴더에서)
    employee_mapper = EmployeeMapper()
    folder_id = await email_client._get_folder_id(settings.target_folder)
    previous_deductions = await fetch_previous_deductions(email_client, folder_id)
    logger.info(f"이전 차감 이력: {len(previous_deductions)}명")

    # 7. 휴가차감 대상 계산 및 메일 발송
    deductions = await send_deduction_emails(
        results=results,
        email_client=email_client,
        employee_mapper=employee_mapper,
        previous_deductions=previous_deductions,
        test_mode=True  # 테스트 모드 - wyyu@envision.co.kr로 발송
    )

    if deductions:
        logger.info(f"휴가차감 메일 {len(deductions)}건 발송 완료")

    # 8. 엑셀 보고서 생성 (이전 차감 이력 포함)
    report_generator = ExcelReportGenerator()
    excel_bytes = report_generator.generate(results, previous_deductions)

    # 8. 보고서 이메일 발송 (주간 차감 이력 포함)
    today_str = datetime.now().strftime("%Y%m%d")
    filename = f"근태_보고서_{today_str}.xlsx"

    email_body = _generate_summary_html(results, deductions, weekly_deductions=previous_deductions)

    success = await email_client.send_email_with_attachment(
        to=settings.report_recipients,
        subject=f"[근태 보고서] {today_str}",
        body=email_body,
        attachment_name=filename,
        attachment_bytes=excel_bytes,
        cc=[REPORT_CC_EMAIL] if REPORT_CC_EMAIL else None
    )

    if success:
        logger.info(f"보고서 발송 완료: {settings.report_recipients}")
    else:
        logger.error("보고서 발송 실패")

    return {
        'success': success,
        'vacations_count': len(results['vacations']),
        'late_arrivals_count': len(results['late_arrivals']),
        'outings_count': len(results['outings']),
        'early_leaves_count': len(results['early_leaves']),
        'unclassified_count': len(results['unclassified']),
        'total_count': len(emails),
        'deductions_count': len(deductions),
        'deductions': deductions
    }


async def send_deduction_emails_only() -> Dict:
    """
    휴가차감 메일만 발송 (평일 오전 7시용)

    Returns:
        Dict: 발송 결과
    """
    logger.info("휴가차감 메일 발송을 시작합니다")

    # 1. Graph API 인증
    auth = GraphAuthenticator(
        tenant_id=settings.azure_tenant_id,
        client_id=settings.azure_client_id,
        client_secret=settings.azure_client_secret
    )
    graph_client = auth.get_client()

    # 2. 이메일 클라이언트 초기화
    email_client = EmailClient(graph_client, settings.target_mailbox)
    classifier = EmailClassifier()
    parser = EmailParser()

    # 3. 폴더 전체 이메일 조회
    emails = await email_client.fetch_emails(
        since=None,
        folder_name=settings.target_folder if settings.target_folder != "Inbox" else None,
        subject_filter="[휴가신고] OR [근태공유]",
        max_count=500
    )

    # 4. 분류 및 파싱
    results = {
        'vacations': [],
        'late_arrivals': [],
        'outings': [],
        'early_leaves': [],
        'unclassified': [],
        'report_date': datetime.now()
    }

    for email in emails:
        await _process_single_email(
            email=email,
            classifier=classifier,
            parser=parser,
            email_client=email_client,
            results=results
        )

    # 5. 이전 휴가차감 이력 조회
    employee_mapper = EmployeeMapper()
    folder_id = await email_client._get_folder_id(settings.target_folder)
    previous_deductions = await fetch_previous_deductions(email_client, folder_id)

    # 6. 휴가차감 메일 발송
    deductions = await send_deduction_emails(
        results=results,
        email_client=email_client,
        employee_mapper=employee_mapper,
        previous_deductions=previous_deductions,
        test_mode=True  # 테스트 모드 - wyyu@envision.co.kr로 발송
    )

    if deductions:
        logger.info(f"휴가차감 메일 {len(deductions)}건 발송 완료")

    return {
        'success': True,
        'deductions_count': len(deductions),
        'deductions': deductions
    }


async def send_report_only() -> Dict:
    """
    보고서만 발송 (평일 오전 8시용)
    차감 메일은 발송하지 않고 보고서만 생성/발송

    Returns:
        Dict: 처리 결과
    """
    logger.info("근태 보고서 발송을 시작합니다")

    # 설정 로드
    settings = Settings.from_env()

    # 1. Graph API 인증
    auth = GraphAuthenticator(
        tenant_id=settings.azure_tenant_id,
        client_id=settings.azure_client_id,
        client_secret=settings.azure_client_secret
    )
    graph_client = auth.get_client()

    # 2. 이메일 클라이언트 초기화
    email_client = EmailClient(graph_client, settings.target_mailbox)
    classifier = EmailClassifier()
    parser = EmailParser()

    # 3. 폴더 전체 이메일 조회
    emails = await email_client.fetch_emails(
        since=None,
        folder_name=settings.target_folder if settings.target_folder != "Inbox" else None,
        subject_filter="[휴가신고] OR [근태공유]",
        max_count=500
    )

    logger.info(f"{len(emails)}개의 이메일을 조회했습니다")

    # 4. 분류 및 파싱
    results = {
        'vacations': [],
        'late_arrivals': [],
        'outings': [],
        'early_leaves': [],
        'unclassified': [],
        'report_date': datetime.now()
    }

    for email in emails:
        await _process_single_email(
            email=email,
            classifier=classifier,
            parser=parser,
            email_client=email_client,
            results=results
        )

    # 5. 로그 출력
    _log_results(results)

    # 6. 이전 휴가차감 이력 조회 (보고서용)
    employee_mapper = EmployeeMapper()
    folder_id = await email_client._get_folder_id(settings.target_folder)
    previous_deductions = await fetch_previous_deductions(email_client, folder_id)

    # 7. 오늘 발송된 차감 메일 정보 수집 (보고서에 표시용)
    today_str = datetime.now().strftime("%Y-%m-%d")
    deductions = []
    for name, history in previous_deductions.items():
        for d in history:
            if d['date'] == today_str:
                deductions.append({
                    'english_name': name,
                    'total_minutes': d['minutes'],  # 근사치
                    'deducted_minutes': d['minutes'],
                    'deduction_days': d['minutes'] / 120 * 0.25
                })

    # 8. 엑셀 보고서 생성
    report_generator = ExcelReportGenerator()
    excel_bytes = report_generator.generate(results, previous_deductions)

    # 9. 이메일 발송 (주간 차감 이력 포함)
    today_str = datetime.now().strftime("%Y%m%d")
    filename = f"근태_보고서_{today_str}.xlsx"

    email_body = _generate_summary_html(results, deductions, weekly_deductions=previous_deductions)

    success = await email_client.send_email_with_attachment(
        to=settings.report_recipients,
        subject=f"[근태 보고서] {today_str}",
        body=email_body,
        attachment_name=filename,
        attachment_bytes=excel_bytes,
        cc=[REPORT_CC_EMAIL] if REPORT_CC_EMAIL else None
    )

    if success:
        logger.info(f"보고서 발송 완료: {settings.report_recipients}")
    else:
        logger.error("보고서 발송 실패")

    return {
        'success': success,
        'vacations_count': len(results['vacations']),
        'late_arrivals_count': len(results['late_arrivals']),
        'outings_count': len(results['outings']),
        'early_leaves_count': len(results['early_leaves']),
        'unclassified_count': len(results['unclassified']),
        'total_count': len(emails)
    }


async def _process_single_email(
    email: EmailMessage,
    classifier: EmailClassifier,
    parser: EmailParser,
    email_client: EmailClient,
    results: Dict
) -> None:
    """단일 이메일 처리"""
    try:
        # 분류
        classification = classifier.classify(email.subject, email.body)

        # 파싱 (제목도 전달하여 휴가 종류 추출)
        extracted = parser.parse(email.body, email.sender_name, email.subject)

        # 부서 정보 조회
        department = await email_client.get_user_department(email.sender_email)

        if classification.category == EmailCategory.VACATION:
            record = VacationRecord(
                applicant=extracted.applicant,
                dates=extracted.dates,
                department=department,
                vacation_type=extracted.vacation_type,
                vacation_days=extracted.vacation_days,
                reason=extracted.reason,
                email_received_at=email.received_at,
                email_id=email.id,
                email_subject=email.subject
            )
            results['vacations'].append(record)

        elif classification.category == EmailCategory.ATTENDANCE:
            record = AttendanceRecord(
                applicant=extracted.applicant,
                sub_type=classification.sub_type,
                date=extracted.dates[0] if extracted.dates else None,
                start_time=extracted.time_range[0] if extracted.time_range else None,
                end_time=extracted.time_range[1] if extracted.time_range else None,
                department=department,
                reason=extracted.reason,
                email_received_at=email.received_at,
                email_id=email.id,
                email_subject=email.subject
            )

            if classification.sub_type == AttendanceSubType.LATE_ARRIVAL:
                results['late_arrivals'].append(record)
            elif classification.sub_type == AttendanceSubType.OUTING:
                results['outings'].append(record)
            elif classification.sub_type == AttendanceSubType.EARLY_LEAVE:
                results['early_leaves'].append(record)
            else:
                results['unclassified'].append(record)

        else:
            logger.warning(f"미분류 이메일: {email.subject}")
            results['unclassified'].append({
                'subject': email.subject,
                'sender': email.sender_name,
                'received_at': email.received_at
            })

    except Exception as e:
        logger.error(f"이메일 처리 중 오류 ({email.subject}): {e}")


def _log_results(results: Dict) -> None:
    """결과 로그 출력"""
    logger.info("=" * 50)
    logger.info("처리 결과:")
    logger.info(f"  - 휴가신고: {len(results['vacations'])}건")
    logger.info(f"  - 출근지연: {len(results['late_arrivals'])}건")
    logger.info(f"  - 외출: {len(results['outings'])}건")
    logger.info(f"  - 조기퇴근: {len(results['early_leaves'])}건")
    logger.info(f"  - 미분류: {len(results['unclassified'])}건")
    logger.info("=" * 50)


def _get_week_dates() -> List[tuple]:
    """이번 주 월요일~금요일 날짜 반환"""
    today = datetime.now()
    # 월요일 찾기 (weekday: 월=0, 화=1, ..., 금=4, 토=5, 일=6)
    days_since_monday = today.weekday()
    monday = today - timedelta(days=days_since_monday)

    week_dates = []
    day_names = ['월', '화', '수', '목', '금']
    for i in range(5):  # 월~금
        date = monday + timedelta(days=i)
        week_dates.append((day_names[i], date.strftime("%Y-%m-%d"), date.strftime("%m/%d")))

    return week_dates


def _generate_summary_html(results: Dict, deductions: List[Dict] = None, weekly_deductions: Dict[str, List[Dict]] = None) -> str:
    """이메일 본문 HTML 생성 (주간 차감 이력 포함)"""
    total = (
        len(results['vacations']) +
        len(results['late_arrivals']) +
        len(results['outings']) +
        len(results['early_leaves'])
    )

    report_date = results.get('report_date', datetime.now())

    # 주간 차감 발송 이력 테이블 생성 (월~금)
    week_dates = _get_week_dates()

    # 날짜별로 차감 데이터 정리
    deductions_by_date = defaultdict(list)
    today_str = datetime.now().strftime("%Y-%m-%d")

    # 이전 차감 이력 추가
    if weekly_deductions:
        for name, history in weekly_deductions.items():
            for d in history:
                deductions_by_date[d['date']].append({
                    'name': name,
                    'minutes': d['minutes']
                })

    # 오늘 발송된 차감 메일 추가
    if deductions:
        for d in deductions:
            deductions_by_date[today_str].append({
                'name': d['english_name'],
                'minutes': d['deducted_minutes']
            })

    # 주간 테이블 행 생성
    weekly_rows = ""
    for day_name, date_full, date_short in week_dates:
        day_deductions = deductions_by_date.get(date_full, [])
        count = len(day_deductions)
        # 이름과 차감시간 함께 표시
        if day_deductions:
            names_with_time = ", ".join([f"{d['name']}({d['minutes']}분)" for d in day_deductions])
        else:
            names_with_time = "-"

        weekly_rows += f"""
            <tr>
                <td>{day_name}</td>
                <td>{date_short}</td>
                <td>{count}건</td>
                <td style="text-align: left; padding-left: 15px;">{names_with_time}</td>
            </tr>"""

    weekly_table = f"""
    <div class="section">
        <h3>주간 차감 발송 이력 (월~금)</h3>
        <table>
            <tr>
                <th style="width: 50px;">요일</th>
                <th style="width: 70px;">날짜</th>
                <th style="width: 60px;">발송</th>
                <th>대상자 (차감시간)</th>
            </tr>
            {weekly_rows}
        </table>
    </div>
    """

    return f"""
    <html>
    <head>
        <style>
            body {{
                font-family: 'Malgun Gothic', Arial, sans-serif;
                font-size: 14px;
                line-height: 1.6;
                color: #333;
                max-width: 500px;
                margin: 0 auto;
                padding: 20px;
            }}
            .header {{
                border-bottom: 2px solid #4472C4;
                padding-bottom: 15px;
                margin-bottom: 25px;
            }}
            .header h2 {{
                color: #4472C4;
                margin: 0 0 10px 0;
                font-size: 22px;
            }}
            .header p {{
                color: #666;
                margin: 0;
                font-size: 13px;
            }}
            .section {{
                margin-bottom: 30px;
            }}
            .section h3 {{
                color: #4472C4;
                font-size: 16px;
                margin-bottom: 15px;
                padding-left: 10px;
                border-left: 4px solid #4472C4;
            }}
            table {{
                width: 100%;
                border-collapse: collapse;
                margin: 15px 0;
                box-shadow: 0 1px 3px rgba(0,0,0,0.1);
            }}
            th {{
                background-color: #4472C4;
                color: white;
                padding: 10px 15px;
                text-align: center;
                font-weight: 500;
            }}
            td {{
                padding: 8px 15px;
                text-align: center;
                border-bottom: 1px solid #eee;
            }}
            tr:nth-child(even) {{
                background-color: #f8f9fa;
            }}
            tr:hover {{
                background-color: #e8f4fc;
            }}
            .total {{
                background-color: #FFF2CC !important;
                font-weight: bold;
            }}
            .total td {{
                border-top: 2px solid #4472C4;
            }}
            .no-data {{
                color: #888;
                font-style: italic;
                padding: 20px;
                background-color: #f8f9fa;
                border-radius: 4px;
                text-align: center;
            }}
            .note {{
                background-color: #e8f4fc;
                padding: 15px 20px;
                border-radius: 4px;
                margin: 20px 0;
                font-size: 13px;
            }}
            .footer {{
                margin-top: 30px;
                padding-top: 20px;
                border-top: 1px solid #eee;
                color: #888;
                font-size: 12px;
                text-align: center;
            }}
        </style>
    </head>
    <body>
        <div class="header">
            <h2>근태 보고서</h2>
            <p>보고서 생성일: {report_date.strftime('%Y년 %m월 %d일 %H:%M')}</p>
        </div>

        {weekly_table}

        <div class="note">
            📎 상세 내역은 첨부된 엑셀 파일을 확인해주세요.
        </div>

        <div class="footer">
            이 메일은 근태 관리 시스템에서 자동으로 발송되었습니다.
        </div>
    </body>
    </html>
    """


def _calculate_minutes(start_time, end_time) -> int:
    """시간 차이를 분으로 계산"""
    if not start_time or not end_time:
        return 0
    start_minutes = start_time.hour * 60 + start_time.minute
    end_minutes = end_time.hour * 60 + end_time.minute
    return max(0, end_minutes - start_minutes)


def _calculate_deduction_days(total_minutes: int) -> float:
    """누적 분을 휴가차감 일수로 변환 (120분 단위)"""
    if total_minutes < 120:
        return 0.0
    # 120분 = 0.25일, 240분 = 0.5일, 360분 = 0.75일, 480분 = 1.0일
    return (total_minutes // 120) * 0.25


def _calculate_deducted_minutes(deduction_days: float) -> int:
    """차감일수에 해당하는 분 계산"""
    return int(deduction_days / 0.25) * 120


async def fetch_previous_deductions(email_client: EmailClient, folder_id: str = None) -> Dict[str, List[Dict]]:
    """
    이전에 발송된 휴가차감 메일에서 차감된 시간을 수집 (근태 자동화 폴더에서)

    Args:
        email_client: 이메일 클라이언트
        folder_id: 폴더 ID (근태 자동화 폴더)

    Returns:
        Dict[str, List[Dict]]: 영문이름 -> [{'date': 날짜, 'minutes': 분}, ...]
    """
    try:
        # 휴가차감 메일 조회 (근태 자동화 폴더에서)
        # 제목 형식: [근태공유] {이름}({일수}일, 휴가차감)
        # 본문에서 시간 추출: 4. 시간: {분}분

        from msgraph.generated.users.item.mail_folders.item.messages.messages_request_builder import MessagesRequestBuilder

        if not folder_id:
            logger.warning("폴더 ID가 제공되지 않았습니다")
            return {}

        # 근태 자동화 폴더에서 메일 조회 (필터 없이 - API InefficientFilter 오류 회피)
        query_params = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            select=["id", "subject", "body", "receivedDateTime"],
            orderby=["receivedDateTime desc"],
            top=500
        )

        request_config = MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )

        response = await email_client.client.users.by_user_id(email_client.mailbox).mail_folders.by_mail_folder_id(folder_id).messages.get(request_configuration=request_config)

        if not response or not response.value:
            logger.info("근태 자동화 폴더에 메일이 없습니다")
            return {}

        # 이름별 차감 이력 수집 (날짜별, 중복 제거)
        deducted_by_name = defaultdict(list)
        seen_deductions = set()  # (이름, 날짜) 조합으로 중복 체크

        for msg in response.value:
            subject = msg.subject or ""

            # Python에서 휴가차감 메일 필터링
            if '휴가차감' not in subject:
                continue

            # 제목에서 이름 추출: [근태공유] English(0.25일, 휴가차감)
            subject_match = re.search(r'\[근태공유\]\s*(\w+)\s*\(', subject)
            if not subject_match:
                continue

            english_name = subject_match.group(1)

            # 본문에서 차감 시간 추출
            body_content = msg.body.content if msg.body else ""
            # HTML에서 텍스트 추출
            body_text = re.sub(r'<[^>]+>', ' ', body_content)

            # "4. 시간: 120분" 또는 "시간:</span> 120분" 패턴
            time_match = re.search(r'시간[:\s]*(\d+)\s*분', body_text)
            if time_match:
                deducted_minutes = int(time_match.group(1))
                # 수신 날짜 추출
                received_date = msg.received_date_time
                if received_date:
                    date_str = received_date.strftime("%Y-%m-%d")
                else:
                    date_str = datetime.now().strftime("%Y-%m-%d")

                # 중복 체크 (같은 사람, 같은 날짜는 한 번만 집계)
                deduction_key = (english_name, date_str)
                if deduction_key in seen_deductions:
                    continue
                seen_deductions.add(deduction_key)

                deducted_by_name[english_name].append({
                    'date': date_str,
                    'minutes': deducted_minutes
                })
                logger.info(f"이전 차감 발견: {english_name} - {date_str} - {deducted_minutes}분")

        logger.info(f"이전 휴가차감 이력: {len(deducted_by_name)}명")
        return dict(deducted_by_name)

    except Exception as e:
        logger.error(f"이전 휴가차감 메일 조회 실패: {e}")
        return {}


async def send_deduction_emails(
    results: Dict,
    email_client: EmailClient,
    employee_mapper: EmployeeMapper,
    previous_deductions: Dict[str, int] = None,
    test_mode: bool = True
) -> List[Dict]:
    """
    누적 120분 이상인 직원에게 휴가차감 메일 발송
    (이전에 발송된 차감 메일의 시간을 제외하고 계산)

    Args:
        results: 처리 결과
        email_client: 이메일 클라이언트
        employee_mapper: 직원 이름 변환기
        previous_deductions: 이전 차감 이력 (영문이름 -> 차감분)
        test_mode: True면 TEST_DEDUCTION_EMAIL로 발송

    Returns:
        발송된 차감 내역 리스트
    """
    if previous_deductions is None:
        previous_deductions = {}

    # 직원별 누적 분 및 근태 내역 수집
    employee_data = defaultdict(lambda: {'minutes': 0, 'records': []})

    all_records = (
        results['late_arrivals'] +
        results['outings'] +
        results['early_leaves']
    )

    for record in all_records:
        minutes = _calculate_minutes(record.start_time, record.end_time)
        employee_data[record.applicant]['minutes'] += minutes
        employee_data[record.applicant]['records'].append(record)

    # 120분 이상인 직원에게 메일 발송 (이전 차감 시간 제외)
    deductions = []
    today_str = datetime.now().strftime("%Y-%m-%d")

    for korean_name, data in employee_data.items():
        total_minutes = data['minutes']
        records = data['records']

        # 영문 이름으로 변환하여 이전 차감 시간 조회
        english_name = employee_mapper.to_english(korean_name)
        deduction_history = previous_deductions.get(english_name, [])
        already_deducted = sum(d['minutes'] for d in deduction_history)

        # 실제 차감 대상 시간 = 총 누적 시간 - 이전에 이미 차감된 시간
        remaining_for_deduction = total_minutes - already_deducted

        logger.info(f"{english_name}: 총 {total_minutes}분, 이전차감 {already_deducted}분, 잔여 {remaining_for_deduction}분")

        # 잔여 시간 기준으로 새로운 차감 계산
        deduction_days = _calculate_deduction_days(remaining_for_deduction)

        if deduction_days > 0:
            deducted_minutes = _calculate_deducted_minutes(deduction_days)

            # 수신자 결정 (테스트 모드면 테스트 주소로)
            recipient = TEST_DEDUCTION_EMAIL if test_mode else f"{english_name.lower()}@envision.co.kr"

            # 메일 제목
            subject = f"[근태공유] {english_name}({deduction_days}일, 휴가차감)"

            # 메일 본문 (근태 내역 포함)
            body = _generate_deduction_html(
                english_name=english_name,
                deduction_days=deduction_days,
                deducted_minutes=deducted_minutes,
                date=today_str,
                records=records
            )

            # 메일 발송 (참조에 관리자 추가)
            success = await email_client.send_email(
                to=[recipient],
                subject=subject,
                body=body,
                cc=[DEDUCTION_CC_EMAIL]
            )

            if success:
                logger.info(f"휴가차감 메일 발송: {english_name} ({deduction_days}일)")
            else:
                logger.error(f"휴가차감 메일 발송 실패: {english_name}")

            deductions.append({
                'korean_name': korean_name,
                'english_name': english_name,
                'total_minutes': total_minutes,
                'deduction_days': deduction_days,
                'deducted_minutes': deducted_minutes,
                'email_sent': success
            })

    return deductions


def _generate_deduction_html(
    english_name: str,
    deduction_days: float,
    deducted_minutes: int,
    date: str,
    records: List[AttendanceRecord] = None
) -> str:
    """휴가차감 메일 HTML 생성 (기존 근태공유 메일 디자인과 동일)"""

    # 근태 내역을 유형별로 분류
    late_arrivals = []
    early_leaves = []
    outings = []

    if records:
        for r in records:
            date_str = r.date.strftime("%Y-%m-%d") if r.date else ""
            minutes = _calculate_minutes(r.start_time, r.end_time) if r.start_time and r.end_time else 0

            if r.sub_type and r.sub_type.value == "출근지연":
                late_arrivals.append((date_str, minutes))
            elif r.sub_type and r.sub_type.value == "조기퇴근":
                early_leaves.append((date_str, minutes))
            elif r.sub_type and r.sub_type.value == "외출":
                outings.append((date_str, minutes))

    # 최대 행 수 계산
    max_rows = max(len(late_arrivals), len(early_leaves), len(outings), 1)

    # 테이블 행 생성
    table_rows = ""
    for i in range(max_rows):
        late_date = late_arrivals[i][0] if i < len(late_arrivals) else ""
        late_min = late_arrivals[i][1] if i < len(late_arrivals) else ""
        early_date = early_leaves[i][0] if i < len(early_leaves) else ""
        early_min = early_leaves[i][1] if i < len(early_leaves) else ""
        out_date = outings[i][0] if i < len(outings) else ""
        out_min = outings[i][1] if i < len(outings) else ""

        table_rows += f"""
            <tr>
                <td>{late_date}</td>
                <td>{late_min}</td>
                <td>{early_date}</td>
                <td>{early_min}</td>
                <td>{out_date}</td>
                <td>{out_min}</td>
            </tr>"""

    return f"""
    <html>
    <head>
        <style>
            body {{
                font-family: 'Malgun Gothic', Arial, sans-serif;
                font-size: 14px;
                line-height: 1.8;
                color: #333;
                max-width: 800px;
                margin: 0 auto;
                padding: 20px;
            }}
            .content {{
                background-color: #fff;
                padding: 30px;
                border-radius: 8px;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            }}
            .item {{
                margin: 15px 0;
            }}
            .label {{
                color: #4472C4;
                font-weight: bold;
            }}
            .history-table {{
                width: 100%;
                border-collapse: collapse;
                margin-top: 15px;
                font-size: 13px;
            }}
            .history-table th {{
                background-color: #4472C4;
                color: white;
                padding: 10px 8px;
                border: 1px solid #ddd;
                text-align: center;
            }}
            .history-table td {{
                padding: 8px;
                border: 1px solid #ddd;
                text-align: center;
            }}
            .footer {{
                margin-top: 30px;
                padding-top: 20px;
                border-top: 1px solid #eee;
                text-align: center;
                color: #666;
                font-size: 12px;
            }}
            .link {{
                color: #4472C4;
                text-decoration: none;
            }}
        </style>
    </head>
    <body>
        <div class="content">
            <p class="item"><span class="label">1. 신고자:</span> {english_name}</p>
            <p class="item"><span class="label">2. 근태공유:</span> 휴가차감</p>
            <p class="item"><span class="label">3. 일자:</span> {date}</p>
            <p class="item"><span class="label">4. 시간:</span> {deducted_minutes}분 ({deduction_days}일)</p>
            <p class="item"><span class="label">5. 근태내역:</span></p>

            <table class="history-table">
                <tr>
                    <th>출근지연-일자</th>
                    <th>출근지연-시간(분)</th>
                    <th>조기퇴근-일자</th>
                    <th>조기퇴근-시간(분)</th>
                    <th>외출-일자</th>
                    <th>외출-시간(분)</th>
                </tr>
                {table_rows}
            </table>
        </div>

        <div class="footer">
            본 메일은 근태 신고 시스템에서 자동으로 발송된 메일입니다.<br><br>
            이 메일은 연차 차감 사전 안내 메일이며, 위 내역과 관련하여 수정이 필요한 부분은 기업발전그룹에 문의 부탁드립니다.<br>
            그룹웨어 상 실제 휴가 차감은 내일 진행될 예정입니다.<br><br>
            <a href="https://attendance-records-375b6.web.app/" class="link">근태 신고 시스템 바로가기</a>
        </div>
    </body>
    </html>
    """
