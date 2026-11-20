---
name: web-design-guidelines
description: Review UI code for Web Interface Guidelines compliance. Use when asked to "review my UI", "check accessibility", "audit design", "review UX", or "check my site against best practices".
metadata:
  author: vercel
  version: "1.0.0"
  argument-hint: <file-or-pattern>
---

# Web Interface Guidelines

Review files for compliance with Web Interface Guidelines.

## How It Works

1. Read local guidelines from `references/rules.md` (or fetch latest from Vercel Labs upstream if network is enabled)
2. Read the specified files or UI codebase
3. Check against all rules in the guidelines: Accessibility, Focus States, Forms, Animation, Layout, Performance
4. Output findings in the terse `file:line` format
