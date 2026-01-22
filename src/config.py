import os
from dataclasses import dataclass
from typing import List
from pathlib import Path

from dotenv import load_dotenv

# .env 파일 로드
env_path = Path(__file__).parent.parent / '.env'
load_dotenv(env_path)


@dataclass
class Settings:
    """애플리케이션 설정"""

    # Azure AD 인증 정보
    azure_tenant_id: str
    azure_client_id: str
    azure_client_secret: str

    # 이메일 설정
    target_mailbox: str
    target_folder: str

    # 보고서 수신자
    report_recipients: List[str]

    # 로깅
    log_level: str

    @classmethod
    def from_env(cls) -> 'Settings':
        """환경 변수에서 설정 로드"""
        recipients_str = os.environ.get('REPORT_RECIPIENTS', '')
        recipients = [r.strip() for r in recipients_str.split(',') if r.strip()]

        return cls(
            azure_tenant_id=os.environ.get('AZURE_TENANT_ID', ''),
            azure_client_id=os.environ.get('AZURE_CLIENT_ID', ''),
            azure_client_secret=os.environ.get('AZURE_CLIENT_SECRET', ''),
            target_mailbox=os.environ.get('TARGET_MAILBOX', ''),
            target_folder=os.environ.get('TARGET_FOLDER', 'Inbox'),
            report_recipients=recipients,
            log_level=os.environ.get('LOG_LEVEL', 'INFO')
        )

    def validate(self) -> None:
        """필수 설정 검증"""
        required = [
            ('AZURE_TENANT_ID', self.azure_tenant_id),
            ('AZURE_CLIENT_ID', self.azure_client_id),
            ('AZURE_CLIENT_SECRET', self.azure_client_secret),
            ('TARGET_MAILBOX', self.target_mailbox),
            ('REPORT_RECIPIENTS', self.report_recipients),
        ]

        missing = [name for name, value in required if not value]

        if missing:
            raise ValueError(f"필수 환경 변수가 설정되지 않았습니다: {', '.join(missing)}")


# 전역 설정 인스턴스
settings = Settings.from_env()
