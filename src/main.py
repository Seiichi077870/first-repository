"""
ピッキングリスト自動生成システム - メインエントリーポイント
"""

import sys
import logging
from pathlib import Path
from datetime import datetime
import pandas as pd

from config import LOG_CONFIG, OUTPUT_DIR
from exceptions import PickingSystemError
from models import ProcessLine, ValidationResult, PickingResult
from utils import (
    setup_logging,
    ensure_directory,
    safe_read_excel,
    safe_write_excel,
    generate_output_filename,
    safe_str,
    log_dataframe_info
)
from processors import (
    ReferenceDBCreator,
    CMPickingProcessor,
    APartsPickingProcessor,
    System720Generator
)

setup_logging(LOG_CONFIG)
logger = logging.getLogger(__name__)


class PickingSystemMain:
    """ピッキングシステムメインクラス"""

    def __init__(self, input_file: Path, process_line: ProcessLine = ProcessLine.A_LINE):
        """
        初期化

        Args:
            input_file: 入力Excelファイルパス
            process_line: 処理ライン
        """
        self.input_file = input_file
        self.process_line = process_line
        self.df_matrix = None
        self.frame_part_number = None
        self.start_time = None

    def run(self) -> PickingResult:
        """
        メイン処理を実行

        Returns:
            処理結果
        """
        self.start_time = datetime.now()

        logger.info("=" * 60)
        logger.info("ピッキングリスト自動生成システム 開始")
        logger.info(f"入力ファイル: {self.input_file}")
        logger.info(f"処理ライン: {self.process_line.value}")
        logger.info("=" * 60)

        try:
            # 入力ファイル検証
            logger.info("--- 入力ファイル検証 ---")
            validation_result = self._validate_input_file()

            if not validation_result.is_valid:
                return PickingResult(
                    success=False,
                    message="入力ファイル検証エラー",
                    errors=validation_result.errors,
                    warnings=validation_result.warnings
                )

            # 参照DB作成
            logger.info("--- 参照DB作成 ---")
            db_creator = ReferenceDBCreator(self.df_matrix)
            df_cm_reference, df_a_reference = db_creator.create_reference_dbs()

            # CMピッキング作成
            logger.info("--- CMピッキング作成 ---")
            cm_processor = CMPickingProcessor(self.df_matrix, df_cm_reference)
            df_cm_picking = cm_processor.create_cm_picking()

            # A部品ピッキング作成
            df_a_picking = None
            df_a_validation = None

            if self.process_line == ProcessLine.A_LINE:
                logger.info("--- A部品ピッキング作成 ---")
                a_processor = APartsPickingProcessor(self.df_matrix, df_a_reference)
                df_a_picking, df_a_validation = a_processor.create_a_parts_picking()

            # 720システム出力作成
            logger.info("--- 720システム出力作成 ---")
            system_720 = System720Generator(
                frame_part_number=self.frame_part_number,
                df_cm_picking=df_cm_picking,
                df_a_parts_picking=df_a_picking,
                df_matrix=self.df_matrix
            )
            wb_720 = system_720.generate()

            # ファイル出力
            output_file = self._save_results(
                df_matrix=self.df_matrix,
                df_cm_reference=df_cm_reference,
                df_cm_picking=df_cm_picking,
                df_a_reference=df_a_reference,
                df_a_picking=df_a_picking,
                df_a_validation=df_a_validation,
                wb_720=wb_720
            )

            processing_time = (datetime.now() - self.start_time).total_seconds()

            logger.info("=" * 60)
            logger.info("ピッキングリスト自動生成完了")
            logger.info(f"出力ファイル: {output_file}")
            logger.info(f"処理時間: {processing_time:.2f}秒")
            logger.info("=" * 60)

            return PickingResult(
                success=True,
                message="処理が正常に完了しました",
                output_file=str(output_file),
                cm_picking_count=len(df_cm_picking),
                a_parts_picking_count=len(df_a_picking) if df_a_picking is not None else 0,
                processing_time=processing_time
            )

        except PickingSystemError as e:
            logger.error(f"処理エラー: {str(e)}")
            return PickingResult(
                success=False,
                message=f"エラー: {str(e)}",
                errors=[str(e)]
            )

        except Exception as e:
            logger.exception("予期しないエラーが発生しました")
            return PickingResult(
                success=False,
                message=f"予期しないエラー: {str(e)}",
                errors=[str(e)]
            )

    def _validate_input_file(self) -> ValidationResult:
        """
        入力ファイルを検証

        Returns:
            検証結果
        """
        errors = []
        warnings = []

        try:
            if not self.input_file.exists():
                errors.append(f"ファイルが存在しません: {self.input_file}")
                return ValidationResult(False, errors, warnings)

            self.df_matrix = safe_read_excel(self.input_file, sheet_name=0, header=None)
            log_dataframe_info(self.df_matrix, "構成表マトリックス")

            if len(self.df_matrix) < 2:
                errors.append("データ行が存在しません")
                return ValidationResult(False, errors, warnings)

            from config import VALIDATION, MATRIX_COLUMNS
            expected_headers = VALIDATION['required_headers']

            for col_idx, expected_value in expected_headers.items():
                if col_idx >= len(self.df_matrix.columns):
                    errors.append(f"列{col_idx + 1}が存在しません")
                    continue

                actual_value = safe_str(self.df_matrix.iloc[0, col_idx])
                if actual_value != expected_value:
                    errors.append(
                        f"列{col_idx + 1}のヘッダーが不正です "
                        f"(期待値: '{expected_value}', 実際: '{actual_value}')"
                    )

            if errors:
                return ValidationResult(False, errors, warnings)

            self.frame_part_number = safe_str(self.df_matrix.iloc[1, MATRIX_COLUMNS['PART_NUMBER']])

            if not self.frame_part_number:
                errors.append("フレーム品番が取得できません")
                return ValidationResult(False, errors, warnings)

            file_prefix = self.input_file.stem[:10] if len(self.input_file.stem) >= 10 else self.input_file.stem
            frame_prefix = self.frame_part_number[:10] if len(self.frame_part_number) >= 10 else self.frame_part_number

            if file_prefix != frame_prefix:
                warnings.append(
                    f"ファイル名とフレーム品番が一致しません "
                    f"(ファイル: '{file_prefix}', フレーム: '{frame_prefix}')"
                )

            logger.info(f"フレーム品番: {self.frame_part_number}")

            for warning in warnings:
                logger.warning(warning)

            return ValidationResult(True, errors, warnings)

        except Exception as e:
            errors.append(f"ファイル検証エラー: {str(e)}")
            return ValidationResult(False, errors, warnings)

    def _save_results(
            self,
            df_matrix: pd.DataFrame,
            df_cm_reference: pd.DataFrame,
            df_cm_picking: pd.DataFrame,
            df_a_reference: pd.DataFrame,
            df_a_picking: pd.DataFrame,
            df_a_validation: pd.DataFrame,
            wb_720
    ) -> Path:
        """
        結果を保存
        """
        ensure_directory(OUTPUT_DIR)
        output_file = generate_output_filename(self.input_file)

        logger.info(f"結果を保存中: {output_file}")

        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            safe_write_excel(writer, df_matrix, "構成表マトリックス", header=False)
            safe_write_excel(writer, df_cm_reference, "CMピッキング参照DB")
            safe_write_excel(writer, df_cm_picking, "CMピッキング")

            if self.process_line == ProcessLine.A_LINE:
                if df_a_reference is not None and not df_a_reference.empty:
                    safe_write_excel(writer, df_a_reference, "A部品ピッキング参照DB")

                if df_a_picking is not None and not df_a_picking.empty:
                    safe_write_excel(writer, df_a_picking, "A部品ピッキング")

                if df_a_validation is not None and not df_a_validation.empty:
                    safe_write_excel(writer, df_a_validation, "A部品ピッキング検証用")

            for sheet in wb_720.worksheets:
                sheet_data = []
                for row in sheet.iter_rows(values_only=True):
                    sheet_data.append(row)

                if sheet_data:
                    df_720 = pd.DataFrame(sheet_data)
                    safe_write_excel(writer, df_720, "720システム入力", header=False)

        logger.info(f"保存完了: {output_file}")
        return output_file


def main():
    """エントリーポイント"""
    import argparse

    parser = argparse.ArgumentParser(
        description="ピッキングリスト自動生成システム",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用例:
  python src/main.py data/input/構成表マトリックス.xlsx
  python src/main.py data/input/構成表マトリックス.xlsx --line C
        """
    )

    parser.add_argument("input_file", type=str, help="入力Excelファイルパス")
    parser.add_argument("--line", choices=["A", "C"], default="A",
                        help="処理ライン (A: CM+A部品, C: CMのみ)")

    args = parser.parse_args()

    input_file = Path(args.input_file)

    if not input_file.exists():
        print(f"❌ エラー: ファイルが見つかりません - {input_file}")
        sys.exit(1)

    process_line = ProcessLine.A_LINE if args.line == "A" else ProcessLine.C_LINE

    system = PickingSystemMain(input_file, process_line)
    result = system.run()

    if result.success:
        print(f"\n✅ {result.message}")
        print(f"   出力ファイル: {result.output_file}")
        print(f"   CMピッキング: {result.cm_picking_count}件")
        if result.a_parts_picking_count > 0:
            print(f"   A部品ピッキング: {result.a_parts_picking_count}件")
        print(f"   処理時間: {result.processing_time:.2f}秒")

        if result.warnings:
            print("\n⚠️  警告:")
            for warning in result.warnings:
                print(f"   - {warning}")

        sys.exit(0)
    else:
        print(f"\n❌ {result.message}")

        if result.errors:
            print("\nエラー詳細:")
            for error in result.errors:
                print(f"   - {error}")

        sys.exit(1)


if __name__ == "__main__":
    main()