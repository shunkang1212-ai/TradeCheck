"""TradeCheck command-line launcher for the local web interface."""

from __future__ import annotations

import argparse
import json
import sys

from . import batch
from .web import server


def main() -> int:
    parser = argparse.ArgumentParser(
        prog="tradecheck",
        description="启动 TradeCheck 网页界面，或批量核对 Excel 单据并输出 JSON",
    )
    parser.add_argument("--port", type=int, default=8753, help="本机端口（默认：8753）")
    commands = parser.add_subparsers(dest="command")
    batch_parser = commands.add_parser("batch", help="按 JSON 清单批量核对单据")
    batch_parser.add_argument("--manifest", required=True, help="批量核对清单 JSON 文件")
    batch_parser.add_argument(
        "--output", default="-", help="JSON 输出文件；默认输出到标准输出（-）"
    )
    batch_parser.add_argument(
        "--force", action="store_true", help="覆盖已存在的 JSON 输出文件，不会覆盖输入文件"
    )
    args = parser.parse_args()

    if args.command == "batch":
        try:
            result, exit_code, protected_paths = batch.run_batch(args.manifest)
            batch.write_report(result, args.output, args.force, protected_paths)
        except (batch.BatchManifestError, OSError) as exc:
            print(f"tradecheck batch: {exc}", file=sys.stderr)
            return 2
        return exit_code

    try:
        httpd = server.run(port=args.port)
    except RuntimeError as exc:
        print(str(exc))
        return 1

    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n已停止。")
    finally:
        server._clear_session()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
