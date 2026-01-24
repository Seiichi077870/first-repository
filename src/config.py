"""
設定ファイル
すべての定数と設定値を集約
"""

from pathlib import Path

# ====================================================================
# プロジェクトルート
# ====================================================================

PROJECT_ROOT = Path(__file__).parent.parent
DATA_DIR = PROJECT_ROOT / "data"

# ====================================================================
# ファイル設定
# ====================================================================

# マスタDBディレクトリ
MASTER_DB_DIR = DATA_DIR / "master"

# CMマスタDB
CM_MASTER_DB = {
    'file': MASTER_DB_DIR / "CMマスタ.xlsx",
    'sheet': 'CMマスタ',
    'columns': {
        'part_number': 'CM品番',
        'box_code': '箱コード',
        'box_name': '箱名称',
        'storage_location': '格納場所'
    }
}

# A部品マスタDB
A_PARTS_MASTER_DB = {
    'file': MASTER_DB_DIR / "A部品マスタ.xlsx",
    'sheet': 'A部品マスタ',
    'columns': {
        'part_number': '部品番号',
        'part_name': '部品名称',
        'storage_location': '格納場所',
        'rack': '棚番',
        'quantity_per_box': '箱入数'
    }
}

# 入力ディレクトリ
INPUT_DIR = DATA_DIR / "input"

# 出力ディレクトリ
OUTPUT_DIR = DATA_DIR / "output"

# ====================================================================
# 列インデックス設定（構成表マトリックス）
# ====================================================================

MATRIX_COLUMNS = {
    'NO': 0,
    'START_PART': 1,
    'FACTORY': 2,
    'PART_NUMBER': 4,
    'PART_NAME': 5,
    'SPEC': 6,
    'COLOR_OUTSIDE': 7,
    'COLOR_INSIDE': 8,
    'QUANTITY': 9,
    'UNIT': 10,
    'FIRST_OPTION': 11
}

# オプション列の開始位置
OPTION_START_COL = 11

# ====================================================================
# 処理対象品番パターン
# ====================================================================

# CM品番パターン（正規表現）
CM_PART_PATTERNS = [
    r'^CM\d{6}',        # CM123456
    r'^C-\d{6}',        # C-123456
]

# A部品品番パターン（正規表現）
A_PARTS_PATTERNS = [
    r'^A\d{6}',         # A123456
    r'^AP-\d{6}',       # AP-123456
]

# 除外品番パターン
EXCLUDE_PART_PATTERNS = [
    r'^FRAME',          # フレーム品番
    r'^SET',            # セット品番
    r'^KIT',            # キット品番
]

# ====================================================================
# 720システム出力設定
# ====================================================================

SYSTEM_720 = {
    'sheet_name': '720システム入力',
    'columns': [
        '伝票番号',
        '行番号',
        '品番',
        '数量',
        '単位',
        '納期',
        '備考'
    ]
}

# ====================================================================
# ログ設定
# ====================================================================

LOG_CONFIG = {
    'version': 1,
    'disable_existing_loggers': False,
    'formatters': {
        'detailed': {
            'format': '%(asctime)s [%(levelname)s] %(name)s: %(message)s',
            'datefmt': '%Y-%m-%d %H:%M:%S'
        },
        'simple': {
            'format': '[%(levelname)s] %(message)s'
        }
    },
    'handlers': {
        'console': {
            'class': 'logging.StreamHandler',
            'level': 'INFO',
            'formatter': 'simple',
            'stream': 'ext://sys.stdout'
        },
        'file': {
            'class': 'logging.FileHandler',
            'level': 'DEBUG',
            'formatter': 'detailed',
            'filename': 'picking_system.log',
            'mode': 'a',
            'encoding': 'utf-8'
        }
    },
    'root': {
        'level': 'DEBUG',
        'handlers': ['console', 'file']
    }
}

# ====================================================================
# 検証設定
# ====================================================================

VALIDATION = {
    'required_headers': {
        0: 'No',
        1: '起点部品番号',
        2: '工場',
        4: '部品番号'
    },
    'max_part_number_length': 20,
    'min_quantity': 0,
    'max_quantity': 9999
}

# ====================================================================
# その他設定
# ====================================================================

# ファイル名の日時フォーマット
DATETIME_FORMAT = '%Y%m%d_%H%M%S'

# Excel設定
EXCEL_CONFIG = {
    'engine': 'openpyxl',
    'date_format': 'YYYY-MM-DD',
    'datetime_format': 'YYYY-MM-DD HH:MM:SS'
}