from dataclasses import dataclass
from datetime import date, time, datetime
from typing import Optional

from mail.email_classifier import AttendanceSubType


@dataclass
class AttendanceRecord:
    """근태 기록 데이터 모델"""

    # 필수 정보
    applicant: str                              # 신청자
    sub_type: AttendanceSubType                 # 유형 (출근지연/외출/조기퇴근)

    # 날짜/시간 정보
    date: Optional[date] = None                 # 날짜
    start_time: Optional[time] = None           # 시작 시간
    end_time: Optional[time] = None             # 종료 시간

    # 추가 정보
    department: Optional[str] = None            # 부서
    reason: str = ""                            # 사유
    email_received_at: Optional[datetime] = None  # 이메일 수신 시간

    # 원본 정보
    email_id: Optional[str] = None              # 원본 이메일 ID
    email_subject: Optional[str] = None         # 원본 이메일 제목

    def __str__(self) -> str:
        date_str = self.date.strftime("%Y-%m-%d") if self.date else "날짜없음"
        time_str = ""
        if self.start_time and self.end_time:
            time_str = f" {self.start_time.strftime('%H:%M')}~{self.end_time.strftime('%H:%M')}"
        return f"[{self.sub_type.value}] {self.applicant} - {date_str}{time_str}"

    def to_dict(self) -> dict:
        """딕셔너리로 변환 (엑셀 출력용)"""
        return {
            '신청자': self.applicant,
            '부서': self.department or '',
            '날짜': self.date.strftime("%Y-%m-%d") if self.date else '',
            '시작시간': self.start_time.strftime("%H:%M") if self.start_time else '',
            '종료시간': self.end_time.strftime("%H:%M") if self.end_time else '',
            '사유': self.reason,
            '이메일수신시간': self.email_received_at.strftime("%Y-%m-%d %H:%M") if self.email_received_at else ''
        }
