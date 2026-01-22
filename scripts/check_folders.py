"""Outlook 폴더 목록 확인 스크립트"""
import asyncio
import sys
sys.path.insert(0, '.')

from src.config import settings
from src.auth import GraphAuthenticator
from msgraph.generated.users.item.mail_folders.item.child_folders.child_folders_request_builder import ChildFoldersRequestBuilder


async def check_folders():
    print("Outlook 폴더 목록 조회 중...")
    print(f"대상 메일함: {settings.target_mailbox}")
    print("=" * 50)

    auth = GraphAuthenticator(
        settings.azure_tenant_id,
        settings.azure_client_id,
        settings.azure_client_secret
    )
    client = auth.get_client()

    # 최상위 폴더 조회
    folders = await client.users.by_user_id(settings.target_mailbox).mail_folders.get()

    if folders and folders.value:
        for folder in folders.value:
            print(f"[{folder.display_name}]")

            # 받은편지함 하위 폴더 조회 (모든 폴더)
            if folder.display_name.lower() in ['inbox', '받은 편지함', '받은편지함']:
                try:
                    # 모든 하위 폴더 가져오기 (top=100)
                    query_params = ChildFoldersRequestBuilder.ChildFoldersRequestBuilderGetQueryParameters(
                        top=100
                    )
                    request_config = ChildFoldersRequestBuilder.ChildFoldersRequestBuilderGetRequestConfiguration(
                        query_parameters=query_params
                    )
                    children = await client.users.by_user_id(settings.target_mailbox).mail_folders.by_mail_folder_id(folder.id).child_folders.get(request_configuration=request_config)
                    if children and children.value:
                        for child in children.value:
                            print(f"  └─ [{child.display_name}]")
                except Exception as e:
                    print(f"  (하위 폴더 조회 실패: {e})")

    print("=" * 50)
    print(f"\n찾고 있는 폴더: [{settings.target_folder}]")


if __name__ == "__main__":
    asyncio.run(check_folders())
