# core/utils.py
# 汎用ユーティリティ

import os

def ensure_dir(path):
    """指定されたパスのディレクトリが存在しない場合は作成する"""
    if not os.path.exists(path):
        os.makedirs(path, exist_ok=True)
