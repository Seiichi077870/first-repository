"""
参照DB作成処理
CMピッキング参照DBとA部品ピッキング参照DBを作成
"""

import logging
from pathlib import Path
from typing import Tuple
import pandas as pd

from config import (
    CM_MASTER_DB,
    A_PARTS_MASTER_DB,
    MATRIX_COLUMNS
)
from exceptions import ReferenceDBError, MasterDBError
from utils import (
    safe_read_excel,
    safe_str,
    identify_part_type,
    is_valid_part_number,
    log_dataframe_info,
    safe_int
)
from models import PartType

logger = logging.getLogger(__name__)


class ReferenceDBCreator:
    """参照DB作成クラス"""

    def __init__(self, df_matrix: pd.DataFrame):
        """
        初期化

        Args:
            df_matrix: 構成表マトリックスDataFrame
        """
        self.df_matrix = df_matrix
        self.df_cm_master = None
        self.df_a_parts_master = None

    def create_reference_dbs(self) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """
        参照DBを作成

        Returns:
            (CMピッキング参照DB, A部品ピッキング参照DB)

        Raises:
            ReferenceDBError: 参照DB作成エラー
        """
        try:
            # マスタDB読み込み
            self._load_master_dbs()

            # CMピッキング参照DB作成
            logger.info("CMピッキング参照DB作成中...")
            df_cm_reference = self._create_cm_reference_db()
            log_dataframe_info(df_cm_reference, "CMピッキング参照DB")

            # A部品ピッキング参照DB作成
            logger.info("A部品ピッキング参照DB作成中...")
            df_a_reference = self._create_a_parts_reference_db()
            log_dataframe_info(df_a_reference, "A部品ピッキング参照DB")

            return df_cm_reference, df_a_reference

        except Exception as e:
            raise ReferenceDBError(f"参照DB作成エラー: {e}")

    def _load_master_dbs(self) -> None:
        """
        マスタDBを読み込み

        Raises:
            MasterDBError: マスタDB読み込みエラー
        """
        try:
            # CMマスタDB
            cm_file = CM_MASTER_DB['file']
            cm_sheet = CM_MASTER_DB['sheet']

            logger.info(f"CMマスタDB読み込み: {cm_file}")
            self.df_cm_master = safe_read_excel(cm_file, sheet_name=cm_sheet, header=0)
            log_dataframe_info(self.df_cm_master, "CMマスタDB")

            # A部品マスタDB
            a_file = A_PARTS_MASTER_DB['file']
            a_sheet = A_PARTS_MASTER_DB['sheet']

            logger.info(f"A部品マスタDB読み込み: {a_file}")
            self.df_a_parts_master = safe_read_excel(a_file, sheet_name=a_sheet, header=0)
            log_dataframe_info(self.df_a_parts_master, "A部品マスタDB")

        except Exception as e:
            raise MasterDBError(f"マスタDB読み込みエラー: {e}")

    def _create_cm_reference_db(self) -> pd.DataFrame:
        """
        CMピッキング参照DBを作成

        Returns:
            CMピッキング参照DataFrame
        """
        cm_reference_data = []

        for idx in range(1, len(self.df_matrix)):
            row = self.df_matrix.iloc[idx]

            part_number = safe_str(row.iloc[MATRIX_COLUMNS['PART_NUMBER']])

            if not is_valid_part_number(part_number):
                continue

            part_type = identify_part_type(part_number)

            if part_type != PartType.CM:
                continue

            master_row = self._find_cm_master(part_number)

            if master_row is None:
                logger.warning(f"CMマスタにない部品: {part_number}")
                continue

            cm_reference_data.append({
                'No': len(cm_reference_data) + 1,
                '起点部品番号': safe_str(row.iloc[MATRIX_COLUMNS['START_PART']]),
                '工場': safe_str(row.iloc[MATRIX_COLUMNS['FACTORY']]),
                '部品番号': part_number,
                '品名': safe_str(row.iloc[MATRIX_COLUMNS['PART_NAME']]),
                '仕様': safe_str(row.iloc[MATRIX_COLUMNS['SPEC']]),
                '箱コード': safe_str(master_row['box_code']),
                '箱名称': safe_str(master_row['box_name']),
                '格納場所': safe_str(master_row['storage_location'])
            })

        df_cm_reference = pd.DataFrame(cm_reference_data)

        if df_cm_reference.empty:
            logger.warning("CM部品が見つかりませんでした")
            df_cm_reference = pd.DataFrame(columns=[
                'No', '起点部品番号', '工場', '部品番号', '品名', '仕様',
                '箱コード', '箱名称', '格納場所'
            ])

        return df_cm_reference

    def _create_a_parts_reference_db(self) -> pd.DataFrame:
        """
        A部品ピッキング参照DBを作成

        Returns:
            A部品ピッキング参照DataFrame
        """
        a_reference_data = []

        for idx in range(1, len(self.df_matrix)):
            row = self.df_matrix.iloc[idx]

            part_number = safe_str(row.iloc[MATRIX_COLUMNS['PART_NUMBER']])

            if not is_valid_part_number(part_number):
                continue

            part_type = identify_part_type(part_number)

            if part_type != PartType.A_PARTS:
                continue

            master_row = self._find_a_parts_master(part_number)

            if master_row is None:
                logger.warning(f"A部品マスタにない部品: {part_number}")
                continue

            a_reference_data.append({
                'No': len(a_reference_data) + 1,
                '起点部品番号': safe_str(row.iloc[MATRIX_COLUMNS['START_PART']]),
                '工場': safe_str(row.iloc[MATRIX_COLUMNS['FACTORY']]),
                '部品番号': part_number,
                '品名': safe_str(row.iloc[MATRIX_COLUMNS['PART_NAME']]),
                '格納場所': safe_str(master_row['storage_location']),
                '棚番': safe_str(master_row['rack']),
                '箱入数': safe_int(master_row['quantity_per_box'])
            })

        df_a_reference = pd.DataFrame(a_reference_data)

        if df_a_reference.empty:
            logger.warning("A部品が見つかりませんでした")
            df_a_reference = pd.DataFrame(columns=[
                'No', '起点部品番号', '工場', '部品番号', '品名',
                '格納場所', '棚番', '箱入数'
            ])

        return df_a_reference

    def _find_cm_master(self, part_number: str) -> dict:
        """
        CMマスタから部品情報を検索

        Args:
            part_number: 部品番号

        Returns:
            マスタ情報辞書（見つからない場合はNone）
        """
        if self.df_cm_master is None:
            return None

        cm_cols = CM_MASTER_DB['columns']
        part_col = cm_cols['part_number']

        mask = self.df_cm_master[part_col] == part_number
        matched = self.df_cm_master[mask]

        if matched.empty:
            return None

        row = matched.iloc[0]

        return {
            'box_code': safe_str(row[cm_cols['box_code']]),
            'box_name': safe_str(row[cm_cols['box_name']]),
            'storage_location': safe_str(row[cm_cols['storage_location']])
        }

    def _find_a_parts_master(self, part_number: str) -> dict:
        """
        A部品マスタから部品情報を検索

        Args:
            part_number: 部品番号

        Returns:
            マスタ情報辞書（見つからない場合はNone）
        """
        if self.df_a_parts_master is None:
            return None

        a_cols = A_PARTS_MASTER_DB['columns']
        part_col = a_cols['part_number']

        mask = self.df_a_parts_master[part_col] == part_number
        matched = self.df_a_parts_master[mask]

        if matched.empty:
            return None

        row = matched.iloc[0]

        return {
            'part_name': safe_str(row[a_cols['part_name']]),
            'storage_location': safe_str(row[a_cols['storage_location']]),
            'rack': safe_str(row[a_cols['rack']]),
            'quantity_per_box': safe_int(row[a_cols['quantity_per_box']])
        }