import logging
from azure.identity import ClientSecretCredential
from msgraph import GraphServiceClient

logger = logging.getLogger(__name__)


class GraphAuthenticator:
    """Microsoft Graph API 인증 클래스"""

    SCOPES = ['https://graph.microsoft.com/.default']

    def __init__(self, tenant_id: str, client_id: str, client_secret: str):
        """
        Graph API 인증 초기화

        Args:
            tenant_id: Azure AD 테넌트 ID
            client_id: 앱 등록 클라이언트 ID
            client_secret: 앱 등록 클라이언트 시크릿
        """
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self._credential = None
        self._client = None

    @property
    def credential(self) -> ClientSecretCredential:
        """Azure 자격 증명 객체 반환"""
        if self._credential is None:
            self._credential = ClientSecretCredential(
                tenant_id=self.tenant_id,
                client_id=self.client_id,
                client_secret=self.client_secret
            )
            logger.info("Azure 자격 증명이 생성되었습니다")
        return self._credential

    def get_client(self) -> GraphServiceClient:
        """
        Graph API 클라이언트 반환

        Returns:
            GraphServiceClient: Microsoft Graph API 클라이언트
        """
        if self._client is None:
            self._client = GraphServiceClient(
                credentials=self.credential,
                scopes=self.SCOPES
            )
            logger.info("Graph API 클라이언트가 생성되었습니다")
        return self._client

    async def verify_connection(self) -> bool:
        """
        Graph API 연결 확인

        Returns:
            bool: 연결 성공 여부
        """
        try:
            client = self.get_client()
            # 간단한 API 호출로 연결 확인
            await client.users.get()
            logger.info("Graph API 연결이 확인되었습니다")
            return True
        except Exception as e:
            logger.error(f"Graph API 연결 실패: {e}")
            return False
