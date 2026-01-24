"""
CMピッキング処理
CMピッキングリストを作成
"""

import logging
from typing import List
import pandas as pd

from config import MATRIX_COLUMNS, OPTION_START_COL
from exceptions import ProcessingError
from utils import (
    safe_str,
    safe_int,
    extract_options,
    format_options,
    log_dataframe_info
)
from models import CMPickingRow

logger = logging.getLogger(__name__)


class CMPickingProcessor:
    """CMピッキング処理クラス"""

    def __init__(self, df_matrix: pd.DataFrame, df_cm_reference: pd.DataFrame):
        """
        初期化

        Args:
            df_matrix: 構成表マトリックスDataFrame
            df_cm_reference: CMピッキング参照DataFrame
        """
        self.df_matrix = df_matrix
        self.df_cm_reference = df_cm_reference

    def create_cm_picking(self) -> pd.DataFrame:
        """
        CMピッキングリストを作成

        Returns:
            CMピッキングDataFrame

        Raises:
            ProcessingError: 処理エラー
        """
        try:
            cm_picking_rows = []

            for idx, ref_row in self.df_cm_reference.iterrows():
                part_number = safe_str(ref_row['部品番号'])

                matrix_row = self._find_matrix_row(part_number)

                if matrix_row is None:
                    logger.warning(f"構成表に見つかりません: {part_number}")
                    continue

                picking_row = self._create_picking_row(ref_row, matrix_row, len(cm_picking_rows) + 1)
                cm_picking_rows.append(picking_row)

            df_cm_picking = self._rows_to_dataframe(cm_picking_rows)

            log_dataframe_info(df_cm_picking, "CMピッキング")
            logger.info(f"CMピッキング作成完了: {len(cm_picking_rows)}件")

            return df_cm_picking

        except Exception as e:
            raise ProcessingError(f"CMピッキング作成エラー: {e}")

    def _find_matrix_row(self, part_number: str) -> pd.Series:
        """
        構成表マトリックスから部品番号に該当する行を検索

        Args:
            part_number: 部品番号

        Returns:
            該当行（見つからない場合はNone）
        """
        for idx in range(1, len(self.df_matrix)):
            row = self.df_matrix.iloc[idx]
            matrix_part_number = safe_str(row.iloc[MATRIX_COLUMNS['PART_NUMBER']])

            if matrix_part_number == part_number:
                return row

        return None

    def _create_picking_row(
            self,
            ref_row: pd.Series,
            matrix_row: pd.Series,
            row_number: int
    ) -> CMPickingRow:
        """
        CMピッキング行を作成

        Args:
            ref_row: 参照DB行
            matrix_row: 構成表マトリックス行
            row_number: 行番号

        Returns:
            CMPickingRow
        """
        start_part_number = safe_str(ref_row['起点部品番号'])
        factory = safe_str(ref_row['工場'])
        part_number = safe_str(ref_row['部品番号'])
        part_name = safe_str(ref_row['品名'])
        spec = safe_str(ref_row['仕様'])

        box_code = safe_str(ref_row['箱コード'])
        box_name = safe_str(ref_row['箱名称'])
        storage_location = safe_str(ref_row['格納場所'])

        quantity = safe_int(matrix_row.iloc[MATRIX_COLUMNS['QUANTITY']])

        options = extract_options(matrix_row, OPTION_START_COL)
        options_str = format_options(options)

        return CMPickingRow(
            no=row_number,
            start_part_number=start_part_number,
            factory=factory,
            part_number=part_number,
            part_name=part_name,
            spec=spec,
            quantity=quantity,
            box_code=box_code,
            box_name=box_name,
            storage_location=storage_location,
            options=options_str
        )

    def _rows_to_dataframe(self, rows: List[CMPickingRow]) -> pd.DataFrame:
        """
        CMPickingRowのリストをDataFrameに変換

        Args:
            rows: CMPickingRowのリスト

        Returns:
            DataFrame
        """
        if not rows:
            return pd.DataFrame(columns=[
                'No', '起点部品番号', '工場', '部品番号', '品名', '仕様',
                '数量', '箱コード', '箱名称', '格納場所', 'オプション'
            ])

        data = [row.to_dict() for row in rows]
        return pd.DataFrame(data)