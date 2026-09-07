"""TradeCheck command-line launcher for the local web interface."""

from __future__ import annotations

import argparse

from .web import server


def main() -> int:
    parser = argparse.ArgumentParser(
        prog="tradecheck",
        description="在本机启动 TradeCheck 外贸单据一致性检查器",
    )
    parser.add_argument("--port", type=int, default=8753, help="本机端口（默认：8753）")
    args = parser.parse_args()

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

