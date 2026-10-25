# compact-reminder の自動お知らせ設定（Claude Code）

スキルだけでもジーニーは長い会話で声をかけますが、容量を正確に測って知らせるには、パソコンの
`~/.claude/settings.json` に次を追加します（すでに中身がある場合は `"hooks"` の部分だけ足す）。

```json
{
  "hooks": {
    "UserPromptSubmit": [
      {
        "hooks": [
          { "type": "command", "command": "python3 ~/.claude/skills/compact-reminder/scripts/context_check.py" }
        ]
      }
    ]
  }
}
```

Windows で `python3` が動かない場合は `python` に変えてください。

- 容量の7割を超えると、ジーニーが返答の最後に `/compact` を勧めます。
- 知らせる割合を変えたいときは環境変数 `COMPACT_REMINDER_THRESHOLD`（例: `0.6`）を設定します。
