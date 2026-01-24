"""
ユーティリティ関数
共通で使用する汎用関数を定義
"""

import re
import logging
from pathlib import Path
from datetime import datetime
from typing import Optional, List, Any, Union
import pandas as pd
from openpyxl import load_workbook, Workbook

from config import (
    CM_PART_PATTERNS,
    A_PARTS_PATTERNS,
    EXCLUDE_PART_PATTERNS,
    DATETIME_FORMAT,
    OUTPUT_DIR
)
from exceptions import (
    FileNotFoundError,
    InvalidFileFormatError
)
from models import PartType

logger = logging.getLogger(__name__)


# ====================================================================
# ファイル操作
# ====================================================================

def ensure_directory(directory: Path) -> None:
    """
    ディレクトリが存在することを保証

    Args:
        directory: ディレクトリパス
    """
    directory.mkdir(parents=True, exist_ok=True)
    logger.debug(f"ディレクトリ確認: {directory}")


def generate_output_filename(input_file: Path, suffix: str = "結果") -> Path:
    """
    出力ファイル名を生成

    Args:
        input_file: 入力ファイルパス
        suffix: ファイル名に付加する接尾辞

    Returns:
        出力ファイルパス
    """
    ensure_directory(OUTPUT_DIR)

    timestamp = datetime.now().strftime(DATETIME_FORMAT)
    base_name = input_file.stem
    output_name = f"{base_name}_{suffix}_{timestamp}.xlsx"

    output_file = OUTPUT_DIR / output_name
    logger.info(f"出力ファイル名: {output_file}")

    return output_file


# ====================================================================
# Excel操作
# ====================================================================

def safe_read_excel(
        file_path: Path,
        sheet_name: Union[str, int] = 0,
        header: Optional[int] = None,
        **kwargs
) -> pd.DataFrame:
    """
    安全にExcelファイルを読み込む

    Args:
        file_path: ファイルパス
        sheet_name: シート名またはインデックス
        header: ヘッダー行（Noneの場合はヘッダーなし）
        **kwargs: pd.read_excelに渡す追加引数

    Returns:
        DataFrame

    Raises:
        FileNotFoundError: ファイルが存在しない
        InvalidFileFormatError: ファイル形式が不正
    """
    if not file_path.exists():
        raise FileNotFoundError(f"ファイルが存在しません: {file_path}")

    try:
        df = pd.read_excel(
            file_path,
            sheet_name=sheet_name,
            header=header,
            engine='openpyxl',
            **kwargs
        )
        logger.debug(f"Excelファイル読み込み成功: {file_path}, Shape: {df.shape}")
        return df

    except Exception as e:
        raise InvalidFileFormatError(f"Excelファイル読み込みエラー: {e}")


def safe_write_excel(
        writer: pd.ExcelWriter,
        df: pd.DataFrame,
        sheet_name: str,
        header: bool = True,
        index: bool = False,
        **kwargs
) -> None:
    """
    安全にDataFrameをExcelに書き込む

    Args:
        writer: ExcelWriter
        df: DataFrame
        sheet_name: シート名
        header: ヘッダーを出力するか
        index: インデックスを出力するか
        **kwargs: to_excelに渡す追加引数
    """
    try:
        df.to_excel(
            writer,
            sheet_name=sheet_name,
            header=header,
            index=index,
            **kwargs
        )
        logger.debug(f"シート書き込み成功: {sheet_name}, Shape: {df.shape}")

    except Exception as e:
        logger.error(f"シート書き込みエラー ({sheet_name}): {e}")
        raise


# ====================================================================
# データ検証
# ====================================================================

def is_valid_part_number(part_number: str) -> bool:
    """
    部品番号の妥当性をチェック

    Args:
        part_number: 部品番号

    Returns:
        妥当な場合True
    """
    if not part_number or not isinstance(part_number, str):
        return False

    part_number = str(part_number).strip()

    if not part_number:
        return False

    if len(part_number) < 6 or len(part_number) > 20:
        return False

    for pattern in EXCLUDE_PART_PATTERNS:
        if re.match(pattern, part_number, re.IGNORECASE):
            return False

    return True


def identify_part_type(part_number: str) -> PartType:
    """
    部品番号から部品タイプを判定

    Args:
        part_number: 部品番号

    Returns:
        部品タイプ
    """
    if not part_number:
        return PartType.OTHER

    part_number = str(part_number).strip()

    for pattern in CM_PART_PATTERNS:
        if re.match(pattern, part_number, re.IGNORECASE):
            return PartType.CM

    for pattern in A_PARTS_PATTERNS:
        if re.match(pattern, part_number, re.IGNORECASE):
            return PartType.A_PARTS

    if re.match(r'^FRAME', part_number, re.IGNORECASE):
        return PartType.FRAME

    return PartType.OTHER


# ====================================================================
# 文字列操作
# ====================================================================

def safe_str(value: Any, default: str = "") -> str:
    """
    安全に文字列に変換

    Args:
        value: 変換する値
        default: 変換できない場合のデフォルト値

    Returns:
        文字列
    """
    if value is None or pd.isna(value):
        return default

    return str(value).strip()


def safe_int(value: Any, default: int = 0) -> int:
    """
    安全に整数に変換

    Args:
        value: 変換する値
        default: 変換できない場合のデフォルト値

    Returns:
        整数
    """
    if value is None or pd.isna(value):
        return default

    try:
        return int(float(value))
    except (ValueError, TypeError):
        return default


def extract_options(row_data: pd.Series, option_start_col: int) -> List[str]:
    """
    行データからオプション情報を抽出

    Args:
        row_data: 行データ
        option_start_col: オプション列の開始位置

    Returns:
        オプションリスト
    """
    options = []

    for col_idx in range(option_start_col, len(row_data)):
        value = safe_str(row_data.iloc[col_idx])
        if value and value not in ['', '-', 'nan']:
            options.append(value)

    return options


def format_options(options: List[str]) -> str:
    """
    オプションリストを文字列に整形

    Args:
        options: オプションリスト

    Returns:
        整形された文字列
    """
    if not options:
        return ""

    return " / ".join(options)


def calculate_required_boxes(quantity: int, quantity_per_box: int) -> int:
    """
    必要箱数を計算

    Args:
        quantity: 必要数量
        quantity_per_box: 箱入数

    Returns:
        必要箱数
    """
    if quantity_per_box <= 0:
        return 0

    return (quantity + quantity_per_box - 1) // quantity_per_box


# ====================================================================
# ロギング
# ====================================================================

def setup_logging(log_config: dict) -> None:
    """
    ロギングを設定

    Args:
        log_config: ログ設定辞書
    """
    import logging.config
    logging.config.dictConfig(log_config)
    logger.info("ロギング設定完了")


def log_dataframe_info(df: pd.DataFrame, name: str = "DataFrame") -> None:
    """
    DataFrameの情報をログ出力

    Args:
        df: DataFrame
        name: DataFrame名
    """
    logger.info(f"{name} - Shape: {df.shape}, Columns: {len(df.columns)}")
    logger.debug(f"{name} - Columns: {df.columns.tolist()}")