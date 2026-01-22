import sys
import os

# Azure Functions에서 모듈을 찾을 수 있도록 경로 추가
current_dir = os.path.dirname(os.path.abspath(__file__))
azure_root = '/home/site/wwwroot'

for path in [azure_root, current_dir]:
    if path not in sys.path:
        sys.path.insert(0, path)

import azure.functions as func
import logging
from datetime import datetime

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

app = func.FunctionApp()


@app.timer_trigger(
    schedule="0 0 18 * * 1-5",  # 평일(월-금) 오후 6시
    arg_name="timer",
    run_on_startup=False,
    use_monitor=True
)
async def deduction_emails(timer: func.TimerRequest) -> None:
    """
    근태차감 메일 발송 (평일 오후 6시)

    누적 근태 시간이 120분 이상인 직원에게 휴가차감 사전 안내 메일 발송
    """
    utc_timestamp = datetime.utcnow().isoformat()

    if timer.past_due:
        logger.warning('타이머가 지연되어 실행되었습니다')

    logger.info(f'근태차감 메일 발송 시작: {utc_timestamp}')

    try:
        from main import send_deduction_emails_only

        result = await send_deduction_emails_only()

        logger.info(f'근태차감 메일 발송 완료: {result["deductions_count"]}건')

    except Exception as e:
        logger.error(f'근태차감 메일 발송 실패: {str(e)}', exc_info=True)
        raise


@app.timer_trigger(
    schedule="0 0 19 * * 1-5",  # 평일(월-금) 오후 7시
    arg_name="timer",
    run_on_startup=False,
    use_monitor=True
)
async def attendance_report(timer: func.TimerRequest) -> None:
    """
    근태 보고서 발송 (평일 오후 7시)

    1. Outlook에서 [휴가신고], [근태공유] 이메일 수집
    2. 이메일 내용 분석 및 분류
    3. 엑셀 보고서 생성
    4. 담당자에게 이메일 발송
    """
    utc_timestamp = datetime.utcnow().isoformat()

    if timer.past_due:
        logger.warning('타이머가 지연되어 실행되었습니다')

    logger.info(f'근태 보고서 발송 시작: {utc_timestamp}')

    try:
        from main import send_report_only

        result = await send_report_only()

        logger.info(f'처리 완료: {result}')

        if result['success']:
            logger.info(
                f"보고서 발송 성공 - "
                f"휴가: {result['vacations_count']}, "
                f"출근지연: {result['late_arrivals_count']}, "
                f"외출: {result['outings_count']}, "
                f"조기퇴근: {result['early_leaves_count']}"
            )
        else:
            logger.error('보고서 발송 실패')

    except Exception as e:
        logger.error(f'보고서 발송 실패: {str(e)}', exc_info=True)
        raise


# HTTP 트리거 (수동 테스트용 - 요청 시에만 실행)
@app.route(route="test/deduction", auth_level=func.AuthLevel.FUNCTION)
async def manual_deduction(req: func.HttpRequest) -> func.HttpResponse:
    """근태차감 메일 수동 테스트"""
    logger.info('근태차감 메일 수동 테스트 요청')

    try:
        from main import send_deduction_emails_only

        result = await send_deduction_emails_only()

        return func.HttpResponse(
            f"근태차감 메일 발송 완료: {result['deductions_count']}건",
            status_code=200
        )

    except Exception as e:
        logger.error(f'수동 테스트 실패: {str(e)}', exc_info=True)
        return func.HttpResponse(f"실행 실패: {str(e)}", status_code=500)


@app.route(route="test/report", auth_level=func.AuthLevel.FUNCTION)
async def manual_report(req: func.HttpRequest) -> func.HttpResponse:
    """보고서 수동 테스트"""
    logger.info('보고서 수동 테스트 요청')

    try:
        from main import send_report_only

        result = await send_report_only()

        return func.HttpResponse(
            f"보고서 발송 완료: 휴가 {result['vacations_count']}건, "
            f"출근지연 {result['late_arrivals_count']}건, "
            f"외출 {result['outings_count']}건, "
            f"조기퇴근 {result['early_leaves_count']}건",
            status_code=200
        )

    except Exception as e:
        logger.error(f'수동 테스트 실패: {str(e)}', exc_info=True)
        return func.HttpResponse(f"실행 실패: {str(e)}", status_code=500)


@app.route(route="test/all", auth_level=func.AuthLevel.FUNCTION)
async def manual_all(req: func.HttpRequest) -> func.HttpResponse:
    """전체 프로세스 수동 테스트 (차감메일 + 보고서)"""
    logger.info('전체 프로세스 수동 테스트 요청')

    try:
        from main import process_attendance_emails

        result = await process_attendance_emails()

        return func.HttpResponse(
            f"전체 처리 완료: 휴가 {result['vacations_count']}건, "
            f"출근지연 {result['late_arrivals_count']}건, "
            f"외출 {result['outings_count']}건, "
            f"조기퇴근 {result['early_leaves_count']}건, "
            f"차감메일 {result['deductions_count']}건",
            status_code=200
        )

    except Exception as e:
        logger.error(f'수동 테스트 실패: {str(e)}', exc_info=True)
        return func.HttpResponse(f"실행 실패: {str(e)}", status_code=500)


# 상태 확인 엔드포인트
@app.route(route="health", auth_level=func.AuthLevel.ANONYMOUS)
def health_check(req: func.HttpRequest) -> func.HttpResponse:
    """헬스 체크 엔드포인트"""
    return func.HttpResponse("OK", status_code=200)
