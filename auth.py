import os
import sys
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from dotenv import load_dotenv
from google.auth.exceptions import RefreshError

# 환경 변수 로드
load_dotenv()

# Google API 접근 범위 설정 (필요한 최소한의 범위만 포함)
SCOPES = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/documents',
    'https://www.googleapis.com/auth/calendar'
]

# 환경 변수에서 클라이언트 정보 가져오기
CLIENT_ID = os.getenv('GOOGLE_CLIENT_ID')
CLIENT_SECRET = os.getenv('GOOGLE_CLIENT_SECRET')

def get_token_path():
    """운영 체제에 따른 토큰 저장 경로 반환"""
    if sys.platform == "win32":
        return os.path.join(os.environ["APPDATA"], "GoogleAPI", "token.json")
    return os.path.join(os.path.expanduser("~"), ".config", "GoogleAPI", "token.json")

def ensure_token_dir():
    """토큰 저장 디렉토리가 없으면 생성"""
    token_dir = os.path.dirname(get_token_path())
    if not os.path.exists(token_dir):
        os.makedirs(token_dir)

def delete_token_file():
    """토큰 파일 삭제"""
    token_path = get_token_path()
    try:
        if os.path.exists(token_path):
            os.remove(token_path)
            print("토큰 파일이 삭제되었습니다.")
    except Exception as e:
        print(f"토큰 파일 삭제 중 오류 발생: {e}")

def get_credentials():
    """OAuth2 인증을 통해 자격 증명 반환"""
    token_path = get_token_path()
    ensure_token_dir()

    creds = None
    # 1. 토큰 파일이 있으면 시도해서 로드
    if os.path.exists(token_path):
        try:
            creds = Credentials.from_authorized_user_file(token_path, SCOPES)
        except Exception as e:
            print(f"토큰 로드 중 오류 발생: {e}")
            delete_token_file()
            creds = None

    # 2. 토큰이 유효하지 않거나 만료되었으면 갱신 시도
    if creds and not creds.valid:
        if creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except RefreshError:
                print("토큰 갱신 실패. 토큰을 삭제하고 재인증을 시도합니다.")
                delete_token_file()
                creds = None
        else:
            delete_token_file()
            creds = None

    # 3. creds가 None이면 새 인증 플로우 시작
    if not creds:
        if not CLIENT_ID or not CLIENT_SECRET:
            raise ValueError("환경 변수에서 GOOGLE_CLIENT_ID와 GOOGLE_CLIENT_SECRET를 찾을 수 없습니다.")
            
        try:
            flow = InstalledAppFlow.from_client_config(
                {
                    "installed": {
                        "client_id": CLIENT_ID,
                        "client_secret": CLIENT_SECRET,
                        "redirect_uris": ["http://localhost:0"],
                        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                        "token_uri": "https://oauth2.googleapis.com/token",
                    }
                },
                SCOPES
            )
            creds = flow.run_local_server(port=0)
            
            # 새 토큰 저장
            with open(token_path, 'w') as token:
                token.write(creds.to_json())
                print("새로운 토큰이 저장되었습니다.")
        except Exception as e:
            print(f"새로운 인증 과정 중 오류 발생: {e}")
            if os.path.exists(token_path):
                delete_token_file()
            raise

    return creds
