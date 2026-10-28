# visio-com

用 **PowerShell + COM** 直接驱动 Microsoft Visio 的 AI Agent Skill。

> 适用于：[Claude Code / Codex / DeepSeek Harness 等](https://github.com) 支持 `SKILL.md` 规约的 Agent。

---

## 安装

把本仓库放进 Agent 的 skills 目录即可：

```powershell
git clone https://github.com/<you>/visio-com.git "$HOME\.agents\skills\visio-diagram"
```

或直接复制目录到 `~/.agents/skills/visio-diagram`。


## 已知限制
- 纯 Windows + 本机 Visio，不能跨平台。