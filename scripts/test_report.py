#!/usr/bin/env python
"""
보고서 발송 테스트 스크립트
"""
import asyncio
import sys
import os

# 프로젝트 루트를 path에 추가
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.main import process_attendance_emails

async def main():
    print("보고서 발송 테스트 시작...")
    result = await process_attendance_emails()
    print(f"발송 결과: {result['success']}")
    print(f"차감 메일: {result['deductions_count']}건")
    print(f"휴가신고: {result['vacations_count']}건")
    print(f"출근지연: {result['late_arrivals_count']}건")
    print(f"외출: {result['outings_count']}건")
    print(f"조기퇴근: {result['early_leaves_count']}건")

if __name__ == "__main__":
    asyncio.run(main())
