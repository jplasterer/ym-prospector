# Starkweather Prospector — Setup Guide

## Folder Structure

Create this structure on your computer before you begin:

```
~/ym-prospector/
├── CLAUDE.md               ← Drop here from downloads
├── ARCHITECTURE.md         ← Drop here from downloads
├── SETUP.md                ← This file
└── output/
    └── seed_list/
        └── starkweather_seed_prospects_v2.xlsx   ← Drop here from downloads
```

Commands to create it:
```bash
mkdir -p ~/ym-prospector/output/seed_list
mkdir ~/ym-prospector/output/nightly
```

---

## Step 1 — Install Node.js (if not already installed)

Claude Code requires Node.js 18 or higher.

Check if you have it:
```bash
node --version
```

If not installed, download from: https://nodejs.org (choose the LTS version)

---

## Step 2 — Install Claude Code

```bash
npm install -g @anthropic-ai/claude-code
```

Verify it installed:
```bash
claude --version
```

---

## Step 3 — Sign In to Claude

```bash
claude
```

On first launch it will open a browser window to authenticate with your
Anthropic/Claude account. Follow the prompts.

---

## Step 4 — Add Web Search (Required for Prospecting)

Claude Code does not have built-in web search. You need to add a web
search MCP server. Choose one:

### Option A — Brave Search (Recommended: Free tier, 2,000 queries/month)

1. Get a free API key at: https://brave.com/search/api
   - Click "Get Started for Free"
   - Select the Free plan (2,000 queries/month)
   - Copy your API key

2. Add the MCP to Claude Code:
```bash
claude mcp add brave-search -e BRAVE_API_KEY=your_api_key_here -- npx -y @modelcontextprotocol/server-brave-search
```

3. Verify it was added:
```bash
claude mcp list
```
You should see `brave-search` in the list.

### Option B — Tavily (More powerful: crawls pages, deeper research)

1. Get a free API key at: https://tavily.com
   - Sign up for a free account
   - Copy your API key from the dashboard

2. Add the MCP to Claude Code:
```bash
claude mcp add tavily -e TAVILY_API_KEY=your_api_key_here -- npx -y tavily-mcp
```

> **Recommendation**: Start with Brave Search (it's simpler and free).
> Add Tavily later if you want Claude to extract full page content from
> association websites during verification.

---

## Step 5 — Open Claude Code in Your Project Folder

```bash
cd ~/ym-prospector
claude
```

Claude Code will automatically read `CLAUDE.md` when it starts. You'll
know it's working because Claude will already know who Starkweather is,
what the ICP is, and what the tool is supposed to do — without you
explaining anything.

---

## Step 6 — Run Your First Prospecting Session

Paste this as your opening prompt (swap in today's date):

```
Run a nightly discovery session per the methodology in CLAUDE.md.

Today's sector focus: Healthcare — nursing, physician assistant,
and pharmacy professional associations in the US.

Target 10–15 candidates. Save output as:
output/nightly/ym_prospects_2026-04-11.xlsx
```

Claude Code will:
1. Search the web for YM associations in that sector
2. Visit each org's website to confirm YM usage
3. Check ProPublica/Candid for revenue and staff
4. Score against the ICP (Tier 1/2/3)
5. Save the Excel file to your output/nightly/ folder

---

## Week-by-Week Sector Rotation

Run one sector per session. Rotate through these in order:

| Week | Sector Focus |
|------|-------------|
| 1 | Healthcare: nursing, physician assistant, pharmacy |
| 2 | Legal: bar associations, paralegal, compliance |
| 3 | Engineering & construction |
| 4 | Finance & accounting: CPA, financial planners |
| 5 | Education & credentialing: HR, training professionals |
| 6 | Technology & IT professionals |
| 7 | Real estate & property management |
| 8 | Nonprofit & association management |
| 9 | Manufacturing & industrial trade |
| 10 | Environmental & energy professionals |
| 11 | Media, marketing & communications |
| 12 | Canada — repeat all sectors with Canada geography filter |

---

## Useful Claude Code Commands

| Command | What it does |
|---------|-------------|
| `claude` | Start a session in the current folder |
| `claude mcp list` | Show all connected MCP tools |
| `claude mcp add ...` | Add a new MCP tool |
| `/help` | Show available slash commands inside a session |
| `/clear` | Clear the conversation context |
| Ctrl+C | Stop the current operation |

---

## Troubleshooting

**"Web search not available"**
→ Your MCP was not added correctly. Run `claude mcp list` to check.
→ Re-run the `claude mcp add` command from Step 4.

**Claude doesn't seem to know the ICP or Starkweather context**
→ Make sure `CLAUDE.md` is in the same folder you ran `claude` from.
→ Run `ls` inside Claude Code to confirm the file is there.

**Excel file not saved**
→ Make sure the `output/nightly/` folder exists: `mkdir -p output/nightly`
→ Ask Claude: "Save the results to output/nightly/ym_prospects_today.xlsx"

**ProPublica lookups are slow**
→ Normal — Claude is visiting each page individually.
→ For faster runs, ask Claude to batch the 990 lookups at the end rather than per-org.

---

## Next Phases (After MVP is Running)

See `ARCHITECTURE.md` for the full roadmap. In brief:

- **Phase 2**: Add Apollo.io plugin for verified decision-maker emails
- **Phase 3**: LinkedIn outreach via Claude in Chrome
- **Phase 4**: Go High Level CRM integration for deduplication and pipeline management
