import re
import logging
from dataclasses import dataclass, field
from datetime import datetime, date, time
from typing import Optional, List, Tuple

logger = logging.getLogger(__name__)


@dataclass
class ExtractedInfo:
    """이메일에서 추출된 정보"""
    applicant: str                          # 신청자
    dates: List[date] = field(default_factory=list)  # 날짜 (복수일 가능)
    reason: str = ""                        # 사유
    time_range: Optional[Tuple[time, time]] = None  # 시간 (시작, 종료)
    vacation_type: Optional[str] = None     # 휴가 종류 (연차, 반차 등)
    vacation_days: Optional[float] = None   # 휴가일수 (이메일에서 추출)
    raw_text: str = ""                      # 원본 텍스트


class EmailParser:
    """이메일 본문에서 정보를 추출하는 파서"""

    # 한국어 날짜 패턴
    DATE_PATTERNS = [
        # 2026 년 1 월 16 일 (공백 포함 형식)
        r'(\d{4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일',
        # 2024년 1월 15일, 2024.1.15, 2024-01-15
        r'(\d{4})[년.\-/]?\s*(\d{1,2})[월.\-/]?\s*(\d{1,2})[일]?',
        # 1월 15일, 1/15
        r'(\d{1,2})[월.\-/]\s*(\d{1,2})[일]?',
    ]

    # 날짜 범위 패턴 (2024.1.15 ~ 2024.1.17)
    DATE_RANGE_PATTERN = r'(\d{4}[년.\-/]?\s*\d{1,2}[월.\-/]?\s*\d{1,2}[일]?)\s*[~\-]\s*(\d{4}[년.\-/]?\s*\d{1,2}[월.\-/]?\s*\d{1,2}[일]?)'

    # 시간 패턴
    TIME_PATTERNS = [
        # 09시 30분 ~ 11시 00분, 09:30 ~ 11:00
        r'(\d{1,2})[시:]\s*(\d{0,2})[분]?\s*[~\-]\s*(\d{1,2})[시:]\s*(\d{0,2})[분]?',
        # 09:30-11:00
        r'(\d{1,2}):(\d{2})\s*[~\-]\s*(\d{1,2}):(\d{2})',
        # 오전 9시 ~ 오후 2시
        r'(오전|오후)\s*(\d{1,2})[시]\s*[~\-]\s*(오전|오후)\s*(\d{1,2})[시]',
    ]

    # 신청자 패턴 (영문 이름 먼저 체크 - 본문에 영문으로 기재된 경우 우선)
    APPLICANT_PATTERNS = [
        r'신고자\s*[:\-]?\s*([A-Za-z]+)',       # 신고자: Janice
        r'신청자\s*[:\-]?\s*([A-Za-z]+)',       # 신청자: Janice
        r'신고자\s*[:\-]?\s*([가-힣]{2,4})',    # 신고자: 홍길동
        r'신청자\s*[:\-]?\s*([가-힣]{2,4})',    # 신청자: 홍길동
        r'성명\s*[:\-]?\s*([가-힣]{2,4})',
        r'이름\s*[:\-]?\s*([가-힣]{2,4})',
        r'작성자\s*[:\-]?\s*([가-힣]{2,4})',
        r'성\s*명\s*[:\-]?\s*([가-힣]{2,4})',
    ]

    # 사유 패턴
    REASON_PATTERNS = [
        r'사유\s*[:\-]?\s*(.+?)(?:\n|$)',
        r'내용\s*[:\-]?\s*(.+?)(?:\n|$)',
        r'비고\s*[:\-]?\s*(.+?)(?:\n|$)',
        r'사\s*유\s*[:\-]?\s*(.+?)(?:\n|$)',
    ]

    # 휴가 종류 패턴
    VACATION_TYPE_PATTERNS = [
        r'(연차|반차|오전반차|오후반차|반반차|병가|경조사|공가|특별휴가|연차휴가|반차휴가)',
        r'휴가\s*종류\s*[:\-]?\s*(\S+)',
        r'휴가종류\s*[:\-]?\s*(\S+)',
    ]

    # 휴가일수 패턴 (이메일에서 직접 추출)
    VACATION_DAYS_PATTERNS = [
        r'휴가\s*일수\s*[:\-]?\s*(\d+(?:\.\d+)?)\s*일',
        r'휴가일수\s*[:\-]?\s*(\d+(?:\.\d+)?)\s*일',
        r'일\s*수\s*[:\-]?\s*(\d+(?:\.\d+)?)\s*일',
        r'일수\s*[:\-]?\s*(\d+(?:\.\d+)?)\s*일',
        r'(\d+(?:\.\d+)?)\s*일\s*(?:사용|신청|휴가)',
        r'총\s*(\d+(?:\.\d+)?)\s*일',
    ]

    def parse(self, body: str, sender_name: str = "", subject: str = "") -> ExtractedInfo:
        """
        이메일 본문에서 정보를 추출합니다.

        Args:
            body: 이메일 본문
            sender_name: 발신자 이름 (신청자 추출 실패 시 사용)
            subject: 이메일 제목 (휴가 종류 추출용)

        Returns:
            ExtractedInfo: 추출된 정보
        """
        # 제목과 본문 모두에서 휴가 종류 추출 시도
        vacation_type = self._extract_vacation_type(subject) or self._extract_vacation_type(body)

        return ExtractedInfo(
            applicant=self._extract_applicant(body, sender_name),
            dates=self._extract_dates(body),
            reason=self._extract_reason(body),
            time_range=self._extract_time_range(body),
            vacation_type=vacation_type,
            vacation_days=self._extract_vacation_days(body),
            raw_text=body
        )

    def _extract_applicant(self, body: str, sender_name: str) -> str:
        """신청자 이름 추출"""
        for pattern in self.APPLICANT_PATTERNS:
            match = re.search(pattern, body)
            if match:
                return match.group(1).strip()

        # 발신자 이름에서 이메일 주소 제거 후 사용
        if sender_name:
            # "홍길동 <hong@company.com>" 형태에서 이름만 추출
            name_match = re.match(r'^([가-힣]+)', sender_name)
            if name_match:
                return name_match.group(1)
            return sender_name.split('<')[0].strip()

        return "미상"

    def _extract_dates(self, body: str) -> List[date]:
        """날짜 추출 (복수일 가능)"""
        dates = []
        current_year = datetime.now().year

        # 전화번호 패턴 제거 (날짜로 오인식 방지)
        clean_body = re.sub(r'0\d{2}[-.\s]?\d{3,4}[-.\s]?\d{4}', '', body)

        # 날짜 범위 먼저 확인
        range_match = re.search(self.DATE_RANGE_PATTERN, clean_body)
        if range_match:
            start_date = self._parse_single_date(range_match.group(1), current_year)
            end_date = self._parse_single_date(range_match.group(2), current_year)
            if start_date and end_date and self._is_reasonable_date(start_date) and self._is_reasonable_date(end_date):
                # 범위 내 모든 날짜 생성 (최대 30일)
                current = start_date
                count = 0
                while current <= end_date and count < 30:
                    dates.append(current)
                    current = self._next_day(current)
                    count += 1

        # 개별 날짜 추출 (날짜 범위 찾은 경우 스킵)
        if not dates:
            for pattern in self.DATE_PATTERNS:
                matches = re.findall(pattern, clean_body)
                for match in matches:
                    try:
                        if len(match) == 3:  # 년월일
                            year = int(match[0])
                            month = int(match[1])
                            day = int(match[2])
                            # 연도가 현재와 ±1년 이내인지 확인
                            if abs(year - current_year) > 1:
                                continue
                            d = date(year, month, day)
                        elif len(match) == 2:  # 월일만
                            month, day = int(match[0]), int(match[1])
                            # 유효한 월/일인지 확인
                            if not (1 <= month <= 12 and 1 <= day <= 31):
                                continue
                            d = date(current_year, month, day)
                        else:
                            continue

                        if d not in dates and self._is_reasonable_date(d):
                            dates.append(d)
                            # 첫 번째 유효한 날짜만 사용 (여러 개 추출 방지)
                            if len(dates) >= 1:
                                break
                    except ValueError:
                        continue
                if dates:
                    break

        return sorted(set(dates)) if dates else []

    def _is_reasonable_date(self, d: date) -> bool:
        """날짜가 합리적인 범위 내인지 확인 (현재 ±1년)"""
        current_year = datetime.now().year
        return abs(d.year - current_year) <= 1

    def _parse_single_date(self, date_str: str, default_year: int) -> Optional[date]:
        """단일 날짜 문자열 파싱"""
        # 년월일 형태
        match = re.search(r'(\d{4})[년.\-/]?\s*(\d{1,2})[월.\-/]?\s*(\d{1,2})', date_str)
        if match:
            try:
                return date(int(match.group(1)), int(match.group(2)), int(match.group(3)))
            except ValueError:
                pass

        # 월일 형태
        match = re.search(r'(\d{1,2})[월.\-/]\s*(\d{1,2})', date_str)
        if match:
            try:
                return date(default_year, int(match.group(1)), int(match.group(2)))
            except ValueError:
                pass

        return None

    def _next_day(self, d: date) -> date:
        """다음 날짜 계산"""
        from datetime import timedelta
        return d + timedelta(days=1)

    def _extract_time_range(self, body: str) -> Optional[Tuple[time, time]]:
        """시간 범위 추출"""
        for pattern in self.TIME_PATTERNS:
            match = re.search(pattern, body)
            if match:
                groups = match.groups()

                # 오전/오후 패턴
                if groups[0] in ('오전', '오후'):
                    start_hour = int(groups[1])
                    if groups[0] == '오후' and start_hour != 12:
                        start_hour += 12
                    end_hour = int(groups[3])
                    if groups[2] == '오후' and end_hour != 12:
                        end_hour += 12
                    return (time(start_hour, 0), time(end_hour, 0))

                # 숫자 시간 패턴
                try:
                    start_hour = int(groups[0])
                    start_min = int(groups[1]) if groups[1] else 0
                    end_hour = int(groups[2])
                    end_min = int(groups[3]) if groups[3] else 0

                    if 0 <= start_hour <= 23 and 0 <= end_hour <= 23:
                        return (time(start_hour, start_min), time(end_hour, end_min))
                except (ValueError, IndexError):
                    continue

        return None

    def _extract_reason(self, body: str) -> str:
        """사유 추출"""
        for pattern in self.REASON_PATTERNS:
            match = re.search(pattern, body, re.MULTILINE)
            if match:
                reason = match.group(1).strip()
                # 너무 긴 사유는 자르기
                if len(reason) > 200:
                    reason = reason[:200] + "..."
                return reason

        return "사유 미기재"

    def _extract_vacation_type(self, body: str) -> Optional[str]:
        """휴가 종류 추출"""
        for pattern in self.VACATION_TYPE_PATTERNS:
            match = re.search(pattern, body)
            if match:
                return match.group(1).strip()

        return None

    def _extract_vacation_days(self, body: str) -> Optional[float]:
        """휴가일수 추출 (이메일에서 직접 추출)"""
        for pattern in self.VACATION_DAYS_PATTERNS:
            match = re.search(pattern, body)
            if match:
                try:
                    days = float(match.group(1))
                    # 유효한 범위인지 확인 (0.25 ~ 30일)
                    if 0.25 <= days <= 30:
                        return days
                except ValueError:
                    continue

        return None
