#!/usr/bin/env python
"""
로컬 테스트 스크립트

Azure AD 연결 없이 파싱/분류/엑셀 생성 기능을 테스트합니다.
"""

import sys
import os
from datetime import datetime, date, time

# 프로젝트 루트를 path에 추가
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.email.email_classifier import EmailClassifier, EmailCategory, AttendanceSubType
from src.email.email_parser import EmailParser
from src.report.excel_generator import ExcelReportGenerator
from src.models.attendance import AttendanceRecord
from src.models.vacation import VacationRecord


def test_classifier():
    """분류기 테스트"""
    print("=" * 50)
    print("이메일 분류기 테스트")
    print("=" * 50)

    classifier = EmailClassifier()

    test_cases = [
        ("[휴가신고] 홍길동 연차 사용", "연차 사용 신청합니다."),
        ("[근태공유] 김철수", "출근지연 안내드립니다. 10시 출근 예정"),
        ("[근태공유] 이영희", "외출 신청합니다. 14:00 ~ 16:00"),
        ("[근태공유] 박민수", "조기퇴근 신청합니다. 17:00"),
        ("일반 이메일", "일반 내용"),
    ]

    for subject, body in test_cases:
        result = classifier.classify(subject, body)
        print(f"제목: {subject}")
        print(f"  → 분류: {result}")
        print()


def test_parser():
    """파서 테스트"""
    print("=" * 50)
    print("이메일 파서 테스트")
    print("=" * 50)

    parser = EmailParser()

    test_bodies = [
        """
        신청자: 홍길동
        날짜: 2024년 1월 15일
        휴가 종류: 연차
        사유: 개인 사정으로 인한 휴가
        """,
        """
        출근지연 안내드립니다.
        성명: 김철수
        날짜: 2024-01-15
        시간: 09:30 ~ 10:00
        사유: 교통 체증
        """,
        """
        외출 신청합니다.
        작성자: 이영희
        날짜: 1월 15일
        시간: 14시 00분 ~ 16시 30분
        사유: 병원 방문
        """,
    ]

    for i, body in enumerate(test_bodies, 1):
        result = parser.parse(body, "발신자")
        print(f"테스트 {i}:")
        print(f"  신청자: {result.applicant}")
        print(f"  날짜: {result.dates}")
        print(f"  시간: {result.time_range}")
        print(f"  사유: {result.reason}")
        print(f"  휴가종류: {result.vacation_type}")
        print()


def test_excel_generator():
    """엑셀 생성기 테스트"""
    print("=" * 50)
    print("엑셀 보고서 생성 테스트")
    print("=" * 50)

    generator = ExcelReportGenerator()

    # 샘플 데이터
    data = {
        'vacations': [
            VacationRecord(
                applicant="홍길동",
                dates=[date(2024, 1, 15)],
                department="개발팀",
                vacation_type="연차",
                reason="개인 사정",
                email_received_at=datetime(2024, 1, 14, 10, 30)
            ),
            VacationRecord(
                applicant="박지민",
                dates=[date(2024, 1, 16), date(2024, 1, 17)],
                department="기획팀",
                vacation_type="반차",
                reason="병원 진료",
                email_received_at=datetime(2024, 1, 15, 9, 0)
            )
        ],
        'late_arrivals': [
            AttendanceRecord(
                applicant="김철수",
                sub_type=AttendanceSubType.LATE_ARRIVAL,
                date=date(2024, 1, 15),
                start_time=time(9, 30),
                end_time=time(10, 0),
                department="영업팀",
                reason="교통 체증",
                email_received_at=datetime(2024, 1, 15, 9, 15)
            )
        ],
        'outings': [
            AttendanceRecord(
                applicant="이영희",
                sub_type=AttendanceSubType.OUTING,
                date=date(2024, 1, 15),
                start_time=time(14, 0),
                end_time=time(16, 0),
                department="마케팅팀",
                reason="거래처 미팅",
                email_received_at=datetime(2024, 1, 15, 13, 30)
            )
        ],
        'early_leaves': [
            AttendanceRecord(
                applicant="최민호",
                sub_type=AttendanceSubType.EARLY_LEAVE,
                date=date(2024, 1, 15),
                start_time=time(17, 0),
                end_time=time(18, 0),
                department="개발팀",
                reason="자녀 병원",
                email_received_at=datetime(2024, 1, 15, 16, 0)
            )
        ],
        'report_date': datetime.now()
    }

    # 엑셀 생성
    excel_bytes = generator.generate(data)

    # 파일로 저장
    output_path = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        f"test_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )

    with open(output_path, 'wb') as f:
        f.write(excel_bytes)

    print(f"엑셀 파일이 생성되었습니다: {output_path}")
    print(f"파일 크기: {len(excel_bytes)} bytes")
    print()
    print("요약:")
    print(f"  - 휴가신고: {len(data['vacations'])}건")
    print(f"  - 출근지연: {len(data['late_arrivals'])}건")
    print(f"  - 외출: {len(data['outings'])}건")
    print(f"  - 조기퇴근: {len(data['early_leaves'])}건")


def main():
    """메인 함수"""
    print("\n근태/휴가 Agent 로컬 테스트\n")

    test_classifier()
    test_parser()
    test_excel_generator()

    print("=" * 50)
    print("모든 테스트가 완료되었습니다!")
    print("=" * 50)


if __name__ == "__main__":
    main()
