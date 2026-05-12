# HTML Build Prompt — AI Tool Interactive Presentation

Paste this prompt into a new Claude Cowork session (same project folder: AI Tool Interactive Presentation).

---

## PROMPT START

You are about to build an interactive single-page HTML website showcasing AI tools for advertising and marketing professionals. Before writing a single line of code, you need to fully understand this project.

**Step 1: Read every file in all three folders.**

This project has three folders with essential context:

- **Context/** — Project overview, writing rules, tone, tool schema, page structure, and two Word documents (AI Philosophy Framework, Tool Positioning Guide)
- **Output/** — Two subfolders: ChatGPT Prompting (contains the updated project brief and a GPT Image 2 ecosystem map prompt with the full color system) and AI Poster (ignore this)
- **Research/** — Three Word documents containing the complete tool cards for all 10 categories. These are the primary content source. Read all three .docx files carefully.

Read every file before doing anything else. The Research folder .docx files are the most critical — they contain the full practitioner-written copy for every tool card that will appear on the site.

**Step 2: Cross-reference the Output folder.**

The file `Output/ChatGPT Prompting/ChatGPT Project Brief — Updated.md` is the authoritative decision log. It confirms final category labels, which tools are included vs. watch list, logo/Clearbit decisions, and any accuracy flags. The file `Output/ChatGPT Prompting/GPT Image 2 — Ecosystem Map Prompt.md` contains the complete per-category color system you will use for the site design.

**Step 3: Ask questions before writing code.**

After reading everything, tell me:
- What you understand the site to be (one paragraph)
- Any gaps, ambiguities, or decisions you need me to make before starting
- Your proposed approach to the two-panel layout interaction pattern

Only start building after I confirm you are good to go.

---

## What You Are Building

A single self-contained HTML file (HTML + CSS + JS, no frameworks, no external dependencies except Clearbit logo API and Google Fonts). The file must work when opened locally and deploy cleanly to GitHub Pages or Netlify.

**Layout and structure:**

- Sticky top navigation bar with anchor links to each of the 10 category sections
- Each category section uses a two-panel layout:
  - Left panel: full expanded tool card for the currently active tool
  - Right panel: scrollable vertical list of all tools in that category; clicking any tool instantly swaps it into the left panel with no page reload
  - The first tool in each category is active by default on load
- Sections flow top to bottom in this order (use these exact labels):
  1. Image Generation
  2. Video Generation
  3. All-in-One Generative AI
  4. AI Agents and Automation
  5. Performance and Paid Media
  6. Social and Creative Tools
  7. Coding and Development
  8. Audio and Voice
  9. Creative Project Management
  10. Research and Insights

**Design direction:**

- Background: #0D1117 (near-black)
- Per-category accent colors (from the GPT Image 2 Ecosystem Map Prompt file — use those exact hex values for each category's active state, section headers, and highlights)
- Clean, dark, editorial feel — no gradients, no gimmicks, no rounded corners on the main panels
- Tool logos pulled from `https://logo.clearbit.com/{domain}` — fallback to a neutral placeholder if logo fails to load
- Typography: Inter (Google Fonts), clean weight hierarchy

**Tool card display (left panel):**

Each expanded tool card shows all fields from the research documents in this order:
- Tool name (large) + company name + logo
- Tagline
- What it is (body copy)
- Best For
- Pros (bulleted)
- Cons (bulleted)
- How to Use It
- Highlight Features
- Pricing
- Link to tool (URL as a styled CTA button)

**Right panel (tool list):**

Each item in the right panel shows: logo (small) + tool name + tagline. Active tool is highlighted with the category accent color. Hover states on all items.

**Watch List:**

The Watch List tools from the Research folder (Category 10 area) get their own section at the bottom, clearly labeled "Watch List," using a card grid layout (not the two-panel format). Each watch list card shows: name, company, what it is (abbreviated), and a link.

---

## Writing and Content Rules

Do not rewrite the tool copy. Use the exact practitioner language from the Research documents. These were written deliberately and the voice matters.

The only writing you may do is:
- UI labels (nav items, button text, section headers, placeholder text)
- Any short UI helper copy not in the research docs

Strictly follow the voice rules in `Context/Tone and Voice.md`: no em dashes, no exclamation points, no buzzwords, lead with why before how.

---

## Technical Requirements

- Single .html file — all CSS and JS inlined
- No frameworks (no React, no Vue, no Bootstrap)
- Google Fonts via `<link>` tag is fine
- Clearbit logo API for tool logos: `https://logo.clearbit.com/{domain}`
- Instant panel swap on click — no page reload, no animation delays over 150ms
- Mobile responsive: on smaller screens, stack panels vertically (tool list on top as a horizontal scroll strip, expanded card below)
- Smooth scroll to sections from nav
- Active nav item updates as user scrolls through sections (IntersectionObserver)
- No localStorage, no cookies, no tracking

---

## Output

Save the completed file to the workspace folder as `AI_Tools_Interactive.html`.

When done, share the file link and tell me: (1) anything you had to make a judgment call on, (2) anything you could not find content for and left as a placeholder, and (3) what you would tackle first in a revision pass.

## PROMPT END
