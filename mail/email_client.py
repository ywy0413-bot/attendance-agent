import logging
import base64
from datetime import datetime, timedelta
from dataclasses import dataclass
from typing import List, Optional

from msgraph import GraphServiceClient
from msgraph.generated.users.item.messages.messages_request_builder import MessagesRequestBuilder
from msgraph.generated.users.item.send_mail.send_mail_post_request_body import SendMailPostRequestBody
from msgraph.generated.users.item.mail_folders.item.child_folders.child_folders_request_builder import ChildFoldersRequestBuilder
from msgraph.generated.models.message import Message
from msgraph.generated.models.item_body import ItemBody
from msgraph.generated.models.body_type import BodyType
from msgraph.generated.models.recipient import Recipient
from msgraph.generated.models.email_address import EmailAddress
from msgraph.generated.models.attachment import Attachment
from msgraph.generated.models.file_attachment import FileAttachment

logger = logging.getLogger(__name__)


@dataclass
class EmailMessage:
    """이메일 메시지 데이터"""
    id: str
    subject: str
    body: str
    sender_name: str
    sender_email: str
    received_at: datetime


class EmailClient:
    """Microsoft Graph API를 사용한 이메일 클라이언트"""

    def __init__(self, graph_client: GraphServiceClient, mailbox: str):
        """
        이메일 클라이언트 초기화

        Args:
            graph_client: Graph API 클라이언트
            mailbox: 대상 메일함 (사용자 이메일 주소)
        """
        self.client = graph_client
        self.mailbox = mailbox

    async def fetch_emails(
        self,
        since: Optional[datetime] = None,
        folder_name: Optional[str] = None,
        subject_filter: Optional[str] = None,
        max_count: int = 100
    ) -> List[EmailMessage]:
        """
        이메일 목록을 가져옵니다.

        Args:
            since: 이 시간 이후의 이메일만 조회
            folder_name: 특정 폴더 (기본값: 받은편지함)
            subject_filter: 제목 필터 (예: "[휴가신고] OR [근태공유]")
            max_count: 최대 조회 수

        Returns:
            List[EmailMessage]: 이메일 목록
        """
        try:
            # 필터 조건 구성
            filter_conditions = []

            if since:
                since_str = since.strftime("%Y-%m-%dT%H:%M:%SZ")
                filter_conditions.append(f"receivedDateTime ge {since_str}")

                # 날짜 필터가 있을 때만 제목 필터 추가 (Graph API 제한 우회)
                if subject_filter:
                    if " OR " in subject_filter:
                        parts = subject_filter.split(" OR ")
                        or_conditions = [f"contains(subject, '{p.strip()}')" for p in parts]
                        filter_conditions.append(f"({' or '.join(or_conditions)})")
                    else:
                        filter_conditions.append(f"contains(subject, '{subject_filter}')")

            filter_query = " and ".join(filter_conditions) if filter_conditions else None

            # 쿼리 파라미터 설정
            query_params = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
                select=["id", "subject", "body", "from", "receivedDateTime"],
                filter=filter_query,
                orderby=["receivedDateTime desc"],
                top=max_count
            )

            request_config = MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
                query_parameters=query_params
            )

            # 폴더 지정 여부에 따라 API 호출
            if folder_name:
                # 폴더 ID 먼저 조회
                folder_id = await self._get_folder_id(folder_name)
                if folder_id:
                    response = await self.client.users.by_user_id(self.mailbox).mail_folders.by_mail_folder_id(folder_id).messages.get(request_configuration=request_config)
                else:
                    logger.warning(f"폴더를 찾을 수 없음: {folder_name}, 받은편지함에서 조회합니다")
                    response = await self.client.users.by_user_id(self.mailbox).messages.get(request_configuration=request_config)
            else:
                response = await self.client.users.by_user_id(self.mailbox).messages.get(request_configuration=request_config)

            if not response or not response.value:
                logger.info("조회된 이메일이 없습니다")
                return []

            # 결과 변환
            emails = []
            for msg in response.value:
                email = EmailMessage(
                    id=msg.id,
                    subject=msg.subject or "",
                    body=self._extract_text_body(msg.body),
                    sender_name=msg.from_.email_address.name if msg.from_ and msg.from_.email_address else "",
                    sender_email=msg.from_.email_address.address if msg.from_ and msg.from_.email_address else "",
                    received_at=msg.received_date_time
                )
                emails.append(email)

            # 날짜 필터 없을 때 Python에서 제목 필터링 적용
            if not since and subject_filter:
                emails = self._filter_by_subject(emails, subject_filter)

            logger.info(f"{len(emails)}개의 이메일을 조회했습니다")
            return emails

        except Exception as e:
            logger.error(f"이메일 조회 실패: {e}")
            raise

    def _filter_by_subject(self, emails: List[EmailMessage], subject_filter: str) -> List[EmailMessage]:
        """제목으로 이메일 필터링 (Python에서 처리)"""
        if " OR " in subject_filter:
            parts = [p.strip() for p in subject_filter.split(" OR ")]
            return [e for e in emails if any(part in e.subject for part in parts)]
        else:
            return [e for e in emails if subject_filter in e.subject]

    async def _get_folder_id(self, folder_name: str) -> Optional[str]:
        """폴더 이름으로 폴더 ID 조회 (하위 폴더 포함)"""
        try:
            # 1. 최상위 폴더에서 찾기
            folders = await self.client.users.by_user_id(self.mailbox).mail_folders.get()
            if folders and folders.value:
                for folder in folders.value:
                    if folder.display_name == folder_name:
                        return folder.id

                # 2. 받은편지함(Inbox) 하위 폴더에서 찾기
                for folder in folders.value:
                    if folder.display_name.lower() in ['inbox', '받은 편지함', '받은편지함']:
                        # 모든 하위 폴더 가져오기 (top=100)
                        query_params = ChildFoldersRequestBuilder.ChildFoldersRequestBuilderGetQueryParameters(
                            top=100
                        )
                        request_config = ChildFoldersRequestBuilder.ChildFoldersRequestBuilderGetRequestConfiguration(
                            query_parameters=query_params
                        )
                        child_folders = await self.client.users.by_user_id(self.mailbox).mail_folders.by_mail_folder_id(folder.id).child_folders.get(request_configuration=request_config)
                        if child_folders and child_folders.value:
                            for child in child_folders.value:
                                if child.display_name == folder_name:
                                    return child.id
            return None
        except Exception as e:
            logger.error(f"폴더 조회 실패: {e}")
            return None

    def _extract_text_body(self, body: Optional[ItemBody]) -> str:
        """이메일 본문에서 텍스트 추출"""
        if not body or not body.content:
            return ""

        content = body.content

        # HTML인 경우 태그 제거
        if body.content_type == BodyType.Html:
            import re
            # HTML 태그 제거
            content = re.sub(r'<[^>]+>', ' ', content)
            # HTML 엔티티 디코딩
            content = content.replace('&nbsp;', ' ')
            content = content.replace('&amp;', '&')
            content = content.replace('&lt;', '<')
            content = content.replace('&gt;', '>')
            content = content.replace('&quot;', '"')
            # 연속 공백 정리
            content = re.sub(r'\s+', ' ', content)

        return content.strip()

    async def get_user_department(self, email: str) -> Optional[str]:
        """
        사용자의 부서 정보 조회

        Args:
            email: 사용자 이메일

        Returns:
            str: 부서명 또는 None
        """
        try:
            user = await self.client.users.by_user_id(email).get()
            if user:
                return user.department
            return None
        except Exception as e:
            logger.debug(f"부서 정보 조회 실패 ({email}): {e}")
            return None

    async def send_email_with_attachment(
        self,
        to: List[str],
        subject: str,
        body: str,
        attachment_name: str,
        attachment_bytes: bytes,
        content_type: str = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ) -> bool:
        """
        첨부파일과 함께 이메일 발송

        Args:
            to: 수신자 이메일 목록
            subject: 이메일 제목
            body: 이메일 본문 (HTML)
            attachment_name: 첨부파일 이름
            attachment_bytes: 첨부파일 바이트 데이터
            content_type: 첨부파일 MIME 타입

        Returns:
            bool: 발송 성공 여부
        """
        try:
            # 수신자 구성
            recipients = [
                Recipient(
                    email_address=EmailAddress(address=addr)
                )
                for addr in to
            ]

            # 첨부파일 구성
            attachment = FileAttachment(
                odata_type="#microsoft.graph.fileAttachment",
                name=attachment_name,
                content_type=content_type,
                content_bytes=attachment_bytes
            )

            # 메시지 구성
            message = Message(
                subject=subject,
                body=ItemBody(
                    content_type=BodyType.Html,
                    content=body
                ),
                to_recipients=recipients,
                attachments=[attachment]
            )

            # 발송 요청
            request_body = SendMailPostRequestBody(
                message=message,
                save_to_sent_items=True
            )

            await self.client.users.by_user_id(self.mailbox).send_mail.post(request_body)

            logger.info(f"이메일 발송 완료: {subject} -> {to}")
            return True

        except Exception as e:
            logger.error(f"이메일 발송 실패: {e}")
            return False

    async def send_email(
        self,
        to: List[str],
        subject: str,
        body: str,
        cc: List[str] = None
    ) -> bool:
        """
        이메일 발송 (첨부파일 없음)

        Args:
            to: 수신자 이메일 목록
            subject: 이메일 제목
            body: 이메일 본문 (HTML)
            cc: 참조 이메일 목록

        Returns:
            bool: 발송 성공 여부
        """
        try:
            # 수신자 구성
            recipients = [
                Recipient(
                    email_address=EmailAddress(address=addr)
                )
                for addr in to
            ]

            # 참조 수신자 구성
            cc_recipients = None
            if cc:
                cc_recipients = [
                    Recipient(
                        email_address=EmailAddress(address=addr)
                    )
                    for addr in cc
                ]

            # 메시지 구성
            message = Message(
                subject=subject,
                body=ItemBody(
                    content_type=BodyType.Html,
                    content=body
                ),
                to_recipients=recipients,
                cc_recipients=cc_recipients
            )

            # 발송 요청
            request_body = SendMailPostRequestBody(
                message=message,
                save_to_sent_items=True
            )

            await self.client.users.by_user_id(self.mailbox).send_mail.post(request_body)

            logger.info(f"이메일 발송 완료: {subject} -> {to}")
            return True

        except Exception as e:
            logger.error(f"이메일 발송 실패: {e}")
            return False
