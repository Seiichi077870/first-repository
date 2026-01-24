"""
A部品ピッキング処理
A部品ピッキングリストを作成
"""

import logging
from typing import List, Tuple
import pandas as pd

from config import MATRIX_COLUMNS, OPTION_START_COL
from exceptions import ProcessingError
from utils import (
    safe_str,
    safe_int,
    extract_options,
    format_options,
    calculate_required_boxes,
    log_dataframe_info
)
from models import APartsPickingRow

logger = logging.getLogger(__name__)


class APartsPickingProcessor:
    """A部品ピッキング処理クラス"""

    def __init__(self, df_matrix: pd.DataFrame, df_a_reference: pd.DataFrame):
        """
        初期化

        Args:
            df_matrix: 構成表マトリックスDataFrame
            df_a_reference: A部品ピッキング参照DataFrame
        """
        self.df_matrix = df_matrix
        self.df_a_reference = df_a_reference

    def create_a_parts_picking(self) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """
        A部品ピッキングリストを作成

        Returns:
            (A部品ピッキングDataFrame, 検証用DataFrame)

        Raises:
            ProcessingError: 処理エラー
        """
        try:
            a_picking_rows = []
            validation_rows = []

            for idx, ref_row in self.df_a_reference.iterrows():
                part_number = safe_str(ref_row['部品番号'])

                matrix_row = self._find_matrix_row(part_number)

                if matrix_row is None:
                    logger.warning(f"構成表に見つかりません: {part_number}")
                    continue

                picking_row = self._create_picking_row(
                    ref_row,
                    matrix_row,
                    len(a_picking_rows) + 1
                )
                a_picking_rows.append(picking_row)

                validation_row = self._create_validation_row(picking_row, matrix_row)
                validation_rows.append(validation_row)

            df_a_picking = self._rows_to_dataframe(a_picking_rows)
            df_validation = self._validation_to_dataframe(validation_rows)

            log_dataframe_info(df_a_picking, "A部品ピッキング")
            logger.info(f"A部品ピッキング作成完了: {len(a_picking_rows)}件")

            return df_a_picking, df_validation

        except Exception as e:
            raise ProcessingError(f"A部品ピッキング作成エラー: {e}")

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
    ) -> APartsPickingRow:
        """
        A部品ピッキング行を作成

        Args:
            ref_row: 参照DB行
            matrix_row: 構成表マトリックス行
            row_number: 行番号

        Returns:
            APartsPickingRow
        """
        start_part_number = safe_str(ref_row['起点部品番号'])
        factory = safe_str(ref_row['工場'])
        part_number = safe_str(ref_row['部品番号'])
        part_name = safe_str(ref_row['品名'])

        storage_location = safe_str(ref_row['格納場所'])
        rack = safe_str(ref_row['棚番'])
        quantity_per_box = safe_int(ref_row['箱入数'])

        quantity = safe_int(matrix_row.iloc[MATRIX_COLUMNS['QUANTITY']])

        required_boxes = calculate_required_boxes(quantity, quantity_per_box)

        options = extract_options(matrix_row, OPTION_START_COL)
        options_str = format_options(options)

        return APartsPickingRow(
            no=row_number,
            start_part_number=start_part_number,
            factory=factory,
            part_number=part_number,
            part_name=part_name,
            quantity=quantity,
            storage_location=storage_location,
            rack=rack,
            quantity_per_box=quantity_per_box,
            required_boxes=required_boxes,
            options=options_str
        )

    def _create_validation_row(
            self,
            picking_row: APartsPickingRow,
            matrix_row: pd.Series
    ) -> dict:
        """
        検証用行を作成（数量チェック用）

        Args:
            picking_row: ピッキング行
            matrix_row: 構成表マトリックス行

        Returns:
            検証用辞書
        """
        calculated_quantity = picking_row.required_boxes * picking_row.quantity_per_box
        original_quantity = picking_row.quantity
        difference = calculated_quantity - original_quantity

        return {
            'No': picking_row.no,
            '部品番号': picking_row.part_number,
            '必要数量': original_quantity,
            '箱入数': picking_row.quantity_per_box,
            '必要箱数': picking_row.required_boxes,
            '実際の数量': calculated_quantity,
            '過不足': difference,
            '判定': 'OK' if difference >= 0 else 'NG'
        }

    def _rows_to_dataframe(self, rows: List[APartsPickingRow]) -> pd.DataFrame:
        """
        APartsPickingRowのリストをDataFrameに変換

        Args:
            rows: APartsPickingRowのリスト

        Returns:
            DataFrame
        """
        if not rows:
            return pd.DataFrame(columns=[
                'No', '起点部品番号', '工場', '部品番号', '品名', '数量',
                '格納場所', '棚番', '箱入数', '必要箱数', 'オプション'
            ])

        data = [row.to_dict() for row in rows]
        return pd.DataFrame(data)

    def _validation_to_dataframe(self, rows: List[dict]) -> pd.DataFrame:
        """
        検証用データをDataFrameに変換

        Args:
            rows: 検証用辞書のリスト

        Returns:
            DataFrame
        """
        if not rows:
            return pd.DataFrame(columns=[
                'No', '部品番号', '必要数量', '箱入数', '必要箱数',
                '実際の数量', '過不足', '判定'
            ])

        return pd.DataFrame(rows)