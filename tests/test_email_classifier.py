import pytest
from src.email.email_classifier import (
    EmailClassifier,
    EmailCategory,
    AttendanceSubType,
    ClassificationResult
)


class TestEmailClassifier:
    """이메일 분류기 테스트"""

    @pytest.fixture
    def classifier(self):
        return EmailClassifier()

    def test_classify_vacation(self, classifier):
        """휴가신고 분류 테스트"""
        subject = "[휴가신고] 홍길동 연차 사용 신청"
        body = "연차 사용 신청합니다. 날짜: 2024년 1월 15일"

        result = classifier.classify(subject, body)

        assert result.category == EmailCategory.VACATION
        assert result.sub_type is None
        assert result.confidence == 1.0

    def test_classify_late_arrival(self, classifier):
        """출근지연 분류 테스트"""
        subject = "[근태공유] 홍길동 출근 지연"
        body = "출근지연 안내드립니다. 10시 출근 예정입니다."

        result = classifier.classify(subject, body)

        assert result.category == EmailCategory.ATTENDANCE
        assert result.sub_type == AttendanceSubType.LATE_ARRIVAL
        assert result.confidence == 0.9

    def test_classify_outing(self, classifier):
        """외출 분류 테스트"""
        subject = "[근태공유] 홍길동"
        body = "외출 신청합니다. 14:00 ~ 16:00 병원 방문"

        result = classifier.classify(subject, body)

        assert result.category == EmailCategory.ATTENDANCE
        assert result.sub_type == AttendanceSubType.OUTING

    def test_classify_early_leave(self, classifier):
        """조기퇴근 분류 테스트"""
        subject = "[근태공유] 홍길동"
        body = "조기퇴근 신청합니다. 17:00 퇴근 예정"

        result = classifier.classify(subject, body)

        assert result.category == EmailCategory.ATTENDANCE
        assert result.sub_type == AttendanceSubType.EARLY_LEAVE

    def test_classify_unknown(self, classifier):
        """미분류 테스트"""
        subject = "일반 이메일 제목"
        body = "일반 이메일 내용"

        result = classifier.classify(subject, body)

        assert result.category == EmailCategory.UNKNOWN
        assert result.confidence == 0.0

    def test_is_target_email(self, classifier):
        """대상 이메일 확인 테스트"""
        assert classifier.is_target_email("[휴가신고] 테스트")
        assert classifier.is_target_email("[근태공유] 테스트")
        assert not classifier.is_target_email("일반 제목")
