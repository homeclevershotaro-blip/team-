#!/usr/bin/env python3
"""UserPromptSubmit フック: 会話の容量が一定を超えたら【容量警告】を伝える。

標準入力の transcript_path から直近の使用トークン数を読み取り、
上限の THRESHOLD 以上なら標準出力にメッセージを出す（Claudeへの追加情報になる）。
"""
import json
import os
import sys

LIMIT = int(os.environ.get("COMPACT_REMINDER_LIMIT", "200000"))
THRESHOLD = float(os.environ.get("COMPACT_REMINDER_THRESHOLD", "0.7"))


def last_usage(path):
    used = 0
    with open(path, encoding="utf-8") as f:
        for line in f:
            try:
                usage = json.loads(line).get("message", {}).get("usage")
            except (ValueError, AttributeError):
                continue
            if isinstance(usage, dict):
                used = (usage.get("input_tokens", 0)
                        + usage.get("cache_read_input_tokens", 0)
                        + usage.get("cache_creation_input_tokens", 0))
    return used


def main():
    try:
        path = json.load(sys.stdin).get("transcript_path", "")
        used = last_usage(path) if path and os.path.exists(path) else 0
    except Exception:
        return
    if used >= LIMIT * THRESHOLD:
        pct = min(99, round(used * 100 / LIMIT))
        print(f"【容量警告】会話の容量が約{pct}%に達しています。"
              "compact-reminder スキルに従い、返答の最後にユーザーへ /compact を勧めてください。")


if __name__ == "__main__":
    main()
