#!/usr/bin/env python3
"""Збирає активність за день із сесій Claude Code + git і друкує Markdown-дайджест.

Детермінований збір «сирого матеріалу» для скіла threads-recap. Далі цей дайджест
читає Claude і на його основі складає варіанти поста для Threads.

Тільки стандартна бібліотека — щоб запускалося будь-де (локально й у веб-контейнері)
без встановлення залежностей.
"""

import argparse
import datetime as dt
import glob
import json
import os
import subprocess
from collections import defaultdict

# Скільки символів лишати від довгих текстів, щоб дайджест був компактним.
PROMPT_MAXLEN = 200
TEXT_MAXLEN = 200
MAX_PROMPTS_PER_GROUP = 30
MAX_TOOLS_SHOWN = 15


def truncate(text, limit):
    text = " ".join(str(text).split())
    return text if len(text) <= limit else text[: limit - 1] + "…"


def default_projects_dir():
    base = os.environ.get("CLAUDE_CONFIG_DIR") or os.path.join(
        os.path.expanduser("~"), ".claude"
    )
    return os.path.join(base, "projects")


def extract_user_text(content):
    """content для user-рядка — рядок або список блоків (tool_result тощо)."""
    if isinstance(content, str):
        return content.strip()
    if isinstance(content, list):
        parts = []
        for block in content:
            if isinstance(block, dict) and block.get("type") == "text":
                parts.append(block.get("text", ""))
        return " ".join(parts).strip()
    return ""


def parse_assistant_content(content):
    """Повертає (список_імен_інструментів, короткий_текст) з assistant-рядка."""
    tools, texts = [], []
    if isinstance(content, list):
        for block in content:
            if not isinstance(block, dict):
                continue
            btype = block.get("type")
            if btype == "tool_use":
                name = block.get("name")
                if name:
                    tools.append(name)
            elif btype == "text":
                t = block.get("text", "").strip()
                if t:
                    texts.append(t)
    return tools, " ".join(texts).strip()


def collect_sessions(projects_dir, date):
    """Проходить */*.jsonl і збирає активність за вказану дату, згруповану за проєктом/гілкою."""
    groups = defaultdict(
        lambda: {"prompts": [], "tools": defaultdict(int), "assistant_texts": []}
    )
    session_ids = set()
    files = glob.glob(os.path.join(projects_dir, "*", "*.jsonl"))

    for path in files:
        try:
            with open(path, "r", encoding="utf-8") as fh:
                lines = fh.readlines()
        except OSError:
            continue

        for line in lines:
            line = line.strip()
            if not line:
                continue
            try:
                obj = json.loads(line)
            except (json.JSONDecodeError, ValueError):
                continue
            if not isinstance(obj, dict):
                continue

            ts = obj.get("timestamp")
            if not (isinstance(ts, str) and ts[:10] == date):
                continue

            ltype = obj.get("type")
            if ltype not in ("user", "assistant"):
                continue

            cwd = obj.get("cwd") or "(невідомий проєкт)"
            branch = obj.get("gitBranch") or "-"
            key = (cwd, branch)
            sid = obj.get("sessionId")
            if sid:
                session_ids.add(sid)

            message = obj.get("message") or {}
            content = message.get("content")

            if ltype == "user":
                # Тільки людські промпти, не tool_result-відповіді від інструментів.
                origin = obj.get("origin") or {}
                is_human = origin.get("kind") == "human" or isinstance(content, str)
                if not is_human:
                    continue
                text = extract_user_text(content)
                if text and not text.startswith("<"):
                    groups[key]["prompts"].append(truncate(text, PROMPT_MAXLEN))
            else:  # assistant
                tools, text = parse_assistant_content(content)
                for name in tools:
                    groups[key]["tools"][name] += 1
                if text:
                    groups[key]["assistant_texts"].append(truncate(text, TEXT_MAXLEN))

    return groups, session_ids


def collect_git(cwd, date):
    """Теми комітів за вказану дату в git-репо cwd. Порожньо, якщо git недоступний."""
    try:
        result = subprocess.run(
            [
                "git",
                "-C",
                cwd,
                "log",
                "--all",
                f'--since={date} 00:00:00',
                f'--until={date} 23:59:59',
                "--pretty=%h|%s",
            ],
            capture_output=True,
            text=True,
            timeout=15,
        )
    except (OSError, subprocess.SubprocessError):
        return []
    if result.returncode != 0:
        return []
    commits = []
    for line in result.stdout.splitlines():
        if "|" in line:
            h, _, subj = line.partition("|")
            commits.append((h.strip(), subj.strip()))
    return commits


def render(groups, session_ids, commits, date):
    out = []
    out.append(f"# Активність за {date}\n")

    if not groups and not commits:
        out.append(
            "_Активності за цю дату не знайдено. Можливо, це нова/порожня сесія, "
            "або скіл запущено у веб-середовищі, де видно лише поточну сесію._\n"
        )
        return "\n".join(out)

    # Секція проєктів
    for (cwd, branch), data in sorted(groups.items()):
        project = os.path.basename(cwd.rstrip("/")) or cwd
        out.append(f"## Проєкт: {project}  (гілка: {branch})")
        out.append(f"_Шлях: {cwd}_\n")

        prompts = data["prompts"][:MAX_PROMPTS_PER_GROUP]
        if prompts:
            out.append("**Мої запити (теми):**")
            for p in prompts:
                out.append(f"- {p}")
            if len(data["prompts"]) > MAX_PROMPTS_PER_GROUP:
                out.append(f"- …ще {len(data['prompts']) - MAX_PROMPTS_PER_GROUP}")
            out.append("")

        tools = data["tools"]
        if tools:
            top = sorted(tools.items(), key=lambda kv: -kv[1])[:MAX_TOOLS_SHOWN]
            summary = ", ".join(f"{name}×{count}" for name, count in top)
            out.append(f"**Що робив Claude (інструменти):** {summary}\n")

    # Секція git
    if commits:
        out.append("## Коміти за день")
        for h, subj in commits:
            out.append(f"- `{h}` {subj}")
        out.append("")

    # Meta
    n_sessions = len(session_ids)
    out.append("## Meta")
    out.append(f"- Сесій за день: {n_sessions}")
    out.append(f"- Проєктів/гілок: {len(groups)}")
    out.append(f"- Комітів: {len(commits)}")
    if n_sessions <= 1:
        out.append(
            "- ⚠️ Видно лише одну сесію — ймовірно, запуск у веб-середовищі "
            "(ефемерний контейнер). Для повної картини дня запускай скіл локально."
        )
    return "\n".join(out)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--date",
        default=dt.date.today().isoformat(),
        help="Дата у форматі YYYY-MM-DD (за замовчуванням — сьогодні за локальним часом).",
    )
    parser.add_argument(
        "--projects-dir",
        default=default_projects_dir(),
        help="Тека з транскриптами сесій Claude Code (за замовч. ~/.claude/projects).",
    )
    args = parser.parse_args()

    groups, session_ids = collect_sessions(args.projects_dir, args.date)
    commits = collect_git(os.getcwd(), args.date)
    print(render(groups, session_ids, commits, args.date))


if __name__ == "__main__":
    main()
