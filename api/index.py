"""Vercel 서버리스 진입점. 루트의 Flask app을 그대로 노출한다."""

import os
import sys

# 루트 디렉터리를 import 경로에 추가 (app.py, naver.py, dart.py 접근)
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app  # noqa: E402  (Vercel의 @vercel/python이 이 `app`을 감지)
