#!/usr/bin/env bash
# 会話の往復を数えて、15往復ごとに /compact を勧める。
# 使い方: compact-reminder.sh count  (UserPromptSubmit)
#         compact-reminder.sh reset  (SessionStart: compact / clear)
input=$(cat)
sid=$(printf '%s' "$input" | sed -n 's/.*"session_id"[[:space:]]*:[[:space:]]*"\([^"]*\)".*/\1/p' | head -n1)
[ -z "$sid" ] && sid=default
dir="${TMPDIR:-/tmp}/claude-compact-reminder"
mkdir -p "$dir"
file="$dir/$sid"

if [ "$1" = "reset" ]; then
  echo 0 > "$file"
  exit 0
fi

n=$(cat "$file" 2>/dev/null || echo 0)
n=$((n + 1))
echo "$n" > "$file"

if [ $((n % 15)) -eq 0 ]; then
  echo "【お知らせ】この会話は${n}往復目です。使える量を節約するため、返答の最後に『会話が長くなってきたので、そろそろ /compact と打ってください（話題が変わるなら /clear）』と一言添えてください。"
fi
exit 0
