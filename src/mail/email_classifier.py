import re
import logging
from enum import Enum
from dataclasses import dataclass
from typing import Optional

logger = logging.getLogger(__name__)


class EmailCategory(Enum):
    """이메일 주요 분류"""
    VACATION = "휴가신고"
    ATTENDANCE = "근태공유"
    UNKNOWN = "미분류"


class AttendanceSubType(Enum):
    """근태공유 하위 분류"""
    LATE_ARRIVAL = "출근지연"
    OUTING = "외출"
    EARLY_LEAVE = "조기퇴근"
    UNKNOWN = "미분류"


@dataclass
class ClassificationResult:
    """분류 결과"""
    category: EmailCategory
    sub_type: Optional[AttendanceSubType]
    confidence: float

    def __str__(self) -> str:
        if self.sub_type:
            return f"{self.category.value} - {self.sub_type.value}"
        return self.category.value


class EmailClassifier:
    """이메일 분류기"""

    # 제목에서 주요 분류 식별 패턴
    VACATION_PATTERN = r'\[휴가신고\]'
    ATTENDANCE_PATTERN = r'\[근태공유\]'

    # 본문에서 근태 하위 분류 식별 패턴
    SUBTYPE_PATTERNS = {
        AttendanceSubType.LATE_ARRIVAL: [
            r'출근\s*지연',
            r'지각',
            r'늦은\s*출근',
            r'출근지연',
        ],
        AttendanceSubType.OUTING: [
            r'외출',
            r'외근',
            r'자리\s*비움',
        ],
        AttendanceSubType.EARLY_LEAVE: [
            r'조기\s*퇴근',
            r'조퇴',
            r'일찍\s*퇴근',
            r'조기퇴근',
        ]
    }

    # 제외할 근태 유형 (취합에서 제외)
    EXCLUDED_PATTERNS = [
        r'당직\s*휴식',
        r'당직휴식',
        r'전일\s*야근',
        r'전일야근',
    ]

    def classify(self, subject: str, body: str) -> ClassificationResult:
        """
        이메일을 분류합니다.

        Args:
            subject: 이메일 제목
            body: 이메일 본문

        Returns:
            ClassificationResult: 분류 결과
        """
        # 1단계: 제목으로 주요 분류
        if re.search(self.VACATION_PATTERN, subject):
            logger.debug(f"휴가신고로 분류됨: {subject}")
            return ClassificationResult(
                category=EmailCategory.VACATION,
                sub_type=None,
                confidence=1.0
            )

        if re.search(self.ATTENDANCE_PATTERN, subject):
            # 제외 유형 확인 (당직휴식, 전일야근)
            if self._is_excluded_type(body):
                logger.debug(f"제외 유형으로 미분류: {subject}")
                return ClassificationResult(
                    category=EmailCategory.UNKNOWN,
                    sub_type=None,
                    confidence=0.0
                )

            # 2단계: 본문에서 하위 분류
            sub_type = self._classify_attendance_subtype(body)
            confidence = 0.9 if sub_type != AttendanceSubType.UNKNOWN else 0.5
            logger.debug(f"근태공유({sub_type.value})로 분류됨: {subject}")
            return ClassificationResult(
                category=EmailCategory.ATTENDANCE,
                sub_type=sub_type,
                confidence=confidence
            )

        logger.warning(f"분류되지 않은 이메일: {subject}")
        return ClassificationResult(
            category=EmailCategory.UNKNOWN,
            sub_type=None,
            confidence=0.0
        )

    def _is_excluded_type(self, body: str) -> bool:
        """
        제외할 근태 유형인지 확인합니다.

        Args:
            body: 이메일 본문

        Returns:
            bool: 제외 대상 여부
        """
        for pattern in self.EXCLUDED_PATTERNS:
            if re.search(pattern, body, re.IGNORECASE):
                return True
        return False

    def _classify_attendance_subtype(self, body: str) -> AttendanceSubType:
        """
        근태공유 이메일의 하위 분류를 결정합니다.

        Args:
            body: 이메일 본문

        Returns:
            AttendanceSubType: 하위 분류 (출근지연/외출/조기퇴근/미분류)
        """
        for subtype, patterns in self.SUBTYPE_PATTERNS.items():
            for pattern in patterns:
                if re.search(pattern, body, re.IGNORECASE):
                    return subtype

        return AttendanceSubType.UNKNOWN

    def is_target_email(self, subject: str) -> bool:
        """
        처리 대상 이메일인지 확인합니다.

        Args:
            subject: 이메일 제목

        Returns:
            bool: 처리 대상 여부
        """
        return bool(
            re.search(self.VACATION_PATTERN, subject) or
            re.search(self.ATTENDANCE_PATTERN, subject)
        )
