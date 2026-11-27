"""ネットワークドライブ（NAS共有）でも同じ操作ができるための補助関数.

ローカルフォルダーでは問題にならない操作が、SMB共有では失敗することがある。

- 「.」始まりの名前を拒否、または隠すサーバーがある（一時フォルダーが消えたように見える）
- 書き出した直後の名前変更が共有違反（WinError 32 など）で一時的に失敗する
- 共有の応答が一時的に止まる、または読み取り専用で書き込めない

名前を作る・動かす・消す操作をこの1箇所に集めて、どの共有先でも同じ結果にする。
名前変更は `os.rename` を経由して呼ぶ（テストは `os.rename` を差し替えて失敗を再現する）。
"""

from __future__ import annotations

import os
import shutil
import tempfile
import time
from pathlib import Path

# 一時フォルダー・一時ファイルの接頭辞。ドット始まりを基本にしつつ、
# 拒否するサーバー向けにチルダ始まりの代替名も同じ一覧へ入れて対象から除外する。
DOT_TEMP_PREFIXES = (".pdf_split_", ".pdf_merge_", ".pdf_rename_")
TILDE_TEMP_PREFIXES = tuple("~" + prefix[1:] for prefix in DOT_TEMP_PREFIXES)
TEMP_NAME_PREFIXES = DOT_TEMP_PREFIXES + TILDE_TEMP_PREFIXES

# 「出力先へコピーして改名」で使う一時ファイルの接頭辞（`.pdf_rename_` を含むので一覧からも除外される）
COPY_TEMP_PREFIX = ".pdf_rename_copy_"

# SMB共有で一時的に起きるエラーだけ再試行する（5=アクセス拒否は再試行しても同じなので除く）
RETRYABLE_WINERRORS = frozenset({32, 33, 64})
RETRY_DELAYS = (0.25, 0.5, 1.0, 2.0)  # 名前変更の再試行間隔（合計 約3.75秒）
CLEANUP_DELAYS = (0.25, 0.5, 1.0)  # 一時フォルダー削除の再試行間隔


def winerror_of(exc: OSError) -> int | None:
    """OSError の Windows エラー番号（無ければ None）."""
    return getattr(exc, "winerror", None)


def is_unc(path: Path) -> bool:
    """ネットワーク上の共有パス（\\\\サーバー\\共有）かどうかを返す."""
    return str(path).replace("/", "\\").startswith("\\\\")


def describe_os_error(exc: Exception, path: Path) -> str:
    """OSエラーに、共有フォルダーで多い原因の案内を添えて返す."""
    lines = [str(exc)]
    if winerror_of(exc) == 5:
        lines.append("アクセスが拒否されました。共有フォルダーが読み取り専用になっていないか、アクセス権を確認してください。")
    elif is_unc(path):
        lines.append("保存先はネットワーク上の共有フォルダーです。共有が一時的に混雑している可能性があります。")
    return "\n".join(lines)


def make_staging_dir(parent: Path, prefix: str) -> Path:
    """出力先フォルダー内に一時フォルダーを作る.

    「~」始まりを先に使う。共有サーバーは「.」始まりを隠す・拒否することがあり、
    途中で止まったときにユーザーが見つけられないため。「~」が使えない共有では
    「.」始まりへ戻る。
    """
    try:
        return Path(tempfile.mkdtemp(prefix=tilde_prefix(prefix), dir=str(parent)))
    except OSError as tilde_error:
        try:
            return Path(tempfile.mkdtemp(prefix=prefix, dir=str(parent)))
        except OSError:
            raise tilde_error from None


def tilde_prefix(prefix: str) -> str:
    """ドット始まりの一時名接頭辞を、チルダ始まりの代替名へ変える."""
    return "~" + prefix.lstrip(".")


def temp_paths_for(folder: Path, prefix: str, count: int) -> list[Path]:
    """フォルダー内で実際に作れる一時名を count 個返す.

    「~」始まりを先に試す。共有サーバーが「.」始まりを隠すと、名前変更の途中で
    止まったときに一時ファイルが見えず、PDFが消えたと誤解されるため。
    「~」が拒否される共有では「.」始まりへ戻る。
    """
    for candidate in (tilde_prefix(prefix), prefix):
        probe = folder / f"{candidate}probe.tmp"
        created = False
        try:
            with probe.open("xb"):
                created = True
            return [folder / f"{candidate}{order}.tmp" for order in range(1, count + 1)]
        except OSError:
            pass
        finally:
            if created:
                remove_file(probe)
    # どちらも試せないときは既定名で進め、実際の操作で出たエラーをそのまま見せる
    return [folder / f"{prefix}{order}.tmp" for order in range(1, count + 1)]


def write_exclusive(path: Path, write_content) -> None:
    """排他作成で書き、内容までサーバーへ届けてから閉じる（SMBでの「書いたのに無い」を防ぐ）."""
    with path.open("xb") as stream:
        write_content(stream)
        stream.flush()
        try:
            os.fsync(stream.fileno())
        except OSError:
            pass  # fsync 非対応の共有でも書き出し自体は完了している


def move_into(source: Path, destination: Path) -> None:
    """名前を動かす。一時的な共有違反は再試行し、それでも動かなければコピー→削除で同じ結果にする.

    既存ファイルへの上書きは行わない。`os.rename` は宛先にファイルがあると
    無条件で上書きするため、先に存在を確認して拒否する（代替のコピーも排他作成）。
    """
    if destination.exists():
        raise FileExistsError(f"宛先に既にファイルがあります。上書きは行いません。\n{destination.name}")
    last_error: OSError | None = None
    for delay in (None,) + RETRY_DELAYS:
        if delay is not None:
            time.sleep(delay)
        try:
            os.rename(source, destination)
            return
        except OSError as exc:
            last_error = exc
            if winerror_of(exc) not in RETRYABLE_WINERRORS:
                raise
    try:
        with source.open("rb") as source_stream, destination.open("xb") as target_stream:
            shutil.copyfileobj(source_stream, target_stream)
    except OSError:
        assert last_error is not None
        raise last_error from None
    remove_file(source)
    if source.exists():
        # 元ファイルを消せなければ「同じ名前が2個」になるため、失敗として扱う
        assert last_error is not None
        raise last_error from None


def copy_into(source: Path, destination: Path) -> None:
    """ファイルをコピーして宛名に置く。既存ファイルへの上書きは行わない.

    共有では「コピーしたのに残っていない」が起きうるため、同じフォルダー内の
    一時名で受け取ってから `move_into` で名前を動かす。元ファイルは変更しない。
    """
    if destination.exists():
        raise FileExistsError(f"宛先に既にファイルがあります。上書きは行いません。\n{destination.name}")
    temp = temp_paths_for(destination.parent, COPY_TEMP_PREFIX, 1)[0]
    try:
        with source.open("rb") as source_stream:
            write_exclusive(temp, lambda stream: shutil.copyfileobj(source_stream, stream))
        move_into(temp, destination)
    finally:
        if temp.exists():
            remove_file(temp)


def remove_file(path: Path) -> None:
    """一時ファイルを消す。無い場合は成功扱い。共有で消えにくいときは少し待ってやり直す."""
    for delay in (None,) + CLEANUP_DELAYS:
        if delay is not None:
            time.sleep(delay)
        try:
            path.unlink()
            return
        except FileNotFoundError:
            return
        except OSError:
            continue


def remove_tree(folder: Path) -> None:
    """一時フォルダーを消す（失敗しても例外は出さない）。共有では少し待ってやり直す."""
    for delay in (None,) + CLEANUP_DELAYS:
        if delay is not None:
            time.sleep(delay)
        shutil.rmtree(folder, ignore_errors=True)
        if not folder.exists():
            return
