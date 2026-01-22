import pytest
from datetime import date, time
from src.email.email_parser import EmailParser, ExtractedInfo


class TestEmailParser:
    """이메일 파서 테스트"""

    @pytest.fixture
    def parser(self):
        return EmailParser()

    def test_extract_applicant_from_body(self, parser):
        """본문에서 신청자 추출 테스트"""
        body = "신청자: 홍길동\n날짜: 2024년 1월 15일"
        result = parser.parse(body, "")

        assert result.applicant == "홍길동"

    def test_extract_applicant_from_sender(self, parser):
        """발신자 이름에서 신청자 추출 테스트"""
        body = "휴가 신청합니다."
        result = parser.parse(body, "김철수")

        assert result.applicant == "김철수"

    def test_extract_date_korean_format(self, parser):
        """한국어 날짜 형식 추출 테스트"""
        body = "휴가 날짜: 2024년 1월 15일"
        result = parser.parse(body, "")

        assert len(result.dates) == 1
        assert result.dates[0] == date(2024, 1, 15)

    def test_extract_date_dash_format(self, parser):
        """대시 날짜 형식 추출 테스트"""
        body = "날짜: 2024-01-15"
        result = parser.parse(body, "")

        assert len(result.dates) == 1
        assert result.dates[0] == date(2024, 1, 15)

    def test_extract_time_range(self, parser):
        """시간 범위 추출 테스트"""
        body = "외출 시간: 14시 00분 ~ 16시 30분"
        result = parser.parse(body, "")

        assert result.time_range is not None
        assert result.time_range[0] == time(14, 0)
        assert result.time_range[1] == time(16, 30)

    def test_extract_time_range_colon_format(self, parser):
        """콜론 시간 형식 추출 테스트"""
        body = "시간: 09:30 ~ 11:00"
        result = parser.parse(body, "")

        assert result.time_range is not None
        assert result.time_range[0] == time(9, 30)
        assert result.time_range[1] == time(11, 0)

    def test_extract_reason(self, parser):
        """사유 추출 테스트"""
        body = "신청자: 홍길동\n사유: 개인 사정으로 인한 연차 사용"
        result = parser.parse(body, "")

        assert result.reason == "개인 사정으로 인한 연차 사용"

    def test_extract_vacation_type(self, parser):
        """휴가 종류 추출 테스트"""
        body = "휴가 종류: 연차\n날짜: 2024년 1월 15일"
        result = parser.parse(body, "")

        assert result.vacation_type == "연차"

    def test_extract_vacation_type_inline(self, parser):
        """인라인 휴가 종류 추출 테스트"""
        body = "반차 사용 신청합니다."
        result = parser.parse(body, "")

        assert result.vacation_type == "반차"
