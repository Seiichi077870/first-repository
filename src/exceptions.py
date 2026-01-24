"""
例外クラス定義
システム固有の例外を定義
"""


class PickingSystemError(Exception):
    """ピッキングシステムの基底例外クラス"""
    pass


class FileNotFoundError(PickingSystemError):
    """ファイルが見つからない"""
    pass


class InvalidFileFormatError(PickingSystemError):
    """ファイルフォーマットが不正"""
    pass


class ValidationError(PickingSystemError):
    """データ検証エラー"""
    pass


class MasterDBError(PickingSystemError):
    """マスタDBエラー"""
    pass


class ProcessingError(PickingSystemError):
    """処理実行エラー"""
    pass


class ReferenceDBError(PickingSystemError):
    """参照DB関連エラー"""
    pass


class OutputError(PickingSystemError):
    """出力処理エラー"""
    pass