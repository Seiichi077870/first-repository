"""
データモデル定義
システムで使用するデータ構造を定義
"""

from dataclasses import dataclass, field
from enum import Enum
from typing import List, Optional, Dict, Any
from pathlib import Path


class ProcessLine(Enum):
    """処理ライン"""
    A_LINE = "A"  # CM + A部品ピッキング
    C_LINE = "C"  # CMピッキングのみ


class PartType(Enum):
    """部品タイプ"""
    CM = "CM"  # CM部品
    A_PARTS = "A_PARTS"  # A部品
    FRAME = "FRAME"  # フレーム
    OTHER = "OTHER"  # その他


@dataclass
class ValidationResult:
    """検証結果"""
    is_valid: bool
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)

    def add_error(self, message: str):
        """エラーを追加"""
        self.errors.append(message)
        self.is_valid = False

    def add_warning(self, message: str):
        """警告を追加"""
        self.warnings.append(message)

    def has_errors(self) -> bool:
        """エラーがあるか"""
        return len(self.errors) > 0

    def has_warnings(self) -> bool:
        """警告があるか"""
        return len(self.warnings) > 0


@dataclass
class CMPickingRow:
    """CMピッキング行"""
    no: int
    start_part_number: str
    factory: str
    part_number: str
    part_name: str
    spec: str
    quantity: int
    box_code: str
    box_name: str
    storage_location: str
    options: str

    def to_dict(self) -> Dict[str, Any]:
        """辞書に変換"""
        return {
            'No': self.no,
            '起点部品番号': self.start_part_number,
            '工場': self.factory,
            '部品番号': self.part_number,
            '品名': self.part_name,
            '仕様': self.spec,
            '数量': self.quantity,
            '箱コード': self.box_code,
            '箱名称': self.box_name,
            '格納場所': self.storage_location,
            'オプション': self.options
        }


@dataclass
class APartsPickingRow:
    """A部品ピッキング行"""
    no: int
    start_part_number: str
    factory: str
    part_number: str
    part_name: str
    quantity: int
    storage_location: str
    rack: str
    quantity_per_box: int
    required_boxes: int
    options: str

    def to_dict(self) -> Dict[str, Any]:
        """辞書に変換"""
        return {
            'No': self.no,
            '起点部品番号': self.start_part_number,
            '工場': self.factory,
            '部品番号': self.part_number,
            '品名': self.part_name,
            '数量': self.quantity,
            '格納場所': self.storage_location,
            '棚番': self.rack,
            '箱入数': self.quantity_per_box,
            '必要箱数': self.required_boxes,
            'オプション': self.options
        }


@dataclass
class System720Row:
    """720システム入力行"""
    slip_number: str  # 伝票番号
    line_number: int  # 行番号
    part_number: str  # 品番
    quantity: int  # 数量
    unit: str  # 単位
    delivery_date: str  # 納期
    remarks: str  # 備考

    def to_list(self) -> List[Any]:
        """リストに変換"""
        return [
            self.slip_number,
            self.line_number,
            self.part_number,
            self.quantity,
            self.unit,
            self.delivery_date,
            self.remarks
        ]


@dataclass
class PickingResult:
    """ピッキング処理結果"""
    success: bool
    message: str
    output_file: Optional[str] = None
    cm_picking_count: int = 0
    a_parts_picking_count: int = 0
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    processing_time: float = 0.0

    def add_error(self, message: str):
        """エラーを追加"""
        self.errors.append(message)
        self.success = False

    def add_warning(self, message: str):
        """警告を追加"""
        self.warnings.append(message)