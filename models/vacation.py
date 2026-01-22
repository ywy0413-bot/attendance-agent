from dataclasses import dataclass
from datetime import date, datetime
from typing import Optional, List


@dataclass
class VacationRecord:
    """휴가 기록 데이터 모델"""

    # 필수 정보
    applicant: str                              # 신청자

    # 날짜 정보
    dates: List[date] = None                    # 휴가 날짜 (복수일 가능)

    # 추가 정보
    department: Optional[str] = None            # 부서
    vacation_type: Optional[str] = None         # 휴가 종류 (연차, 반차 등)
    vacation_days: Optional[float] = None       # 휴가일수 (이메일에서 추출)
    reason: str = ""                            # 사유
    email_received_at: Optional[datetime] = None  # 이메일 수신 시간

    # 원본 정보
    email_id: Optional[str] = None              # 원본 이메일 ID
    email_subject: Optional[str] = None         # 원본 이메일 제목

    def __post_init__(self):
        if self.dates is None:
            self.dates = []

    @property
    def date(self) -> Optional[date]:
        """첫 번째 휴가일 반환 (호환성용)"""
        return self.dates[0] if self.dates else None

    @property
    def date_range_str(self) -> str:
        """날짜 범위 문자열"""
        if not self.dates:
            return "날짜없음"
        if len(self.dates) == 1:
            return self.dates[0].strftime("%Y-%m-%d")
        return f"{self.dates[0].strftime('%Y-%m-%d')} ~ {self.dates[-1].strftime('%Y-%m-%d')}"

    def __str__(self) -> str:
        vtype = f"({self.vacation_type})" if self.vacation_type else ""
        return f"[휴가신고{vtype}] {self.applicant} - {self.date_range_str}"

    def to_dict(self) -> dict:
        """딕셔너리로 변환 (엑셀 출력용)"""
        return {
            '신청자': self.applicant,
            '부서': self.department or '',
            '휴가일자': self.date_range_str,
            '휴가종류': self.vacation_type or '',
            '사유': self.reason,
            '이메일수신시간': self.email_received_at.strftime("%Y-%m-%d %H:%M") if self.email_received_at else ''
        }
