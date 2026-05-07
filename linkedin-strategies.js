const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, LevelFormat, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak
} = require('docx');
const fs = require('fs');

const BRAND_BLUE = "1B4F9B";
const BRAND_LIGHT = "E8F0FC";
const ACCENT_ORANGE = "E07B30";
const GRAY = "555555";
const LIGHT_GRAY = "F5F5F5";
const WHITE = "FFFFFF";
const DARK = "1A1A1A";

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NIL };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 480, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: BRAND_BLUE, space: 4 } },
    children: [new TextRun({ text, font: "Arial", size: 30, bold: true, color: BRAND_BLUE })]
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 320, after: 120 },
    children: [new TextRun({ text, font: "Arial", size: 24, bold: true, color: BRAND_BLUE })]
  });
}

function h3(text) {
  return new Paragraph({
    spacing: { before: 240, after: 80 },
    children: [new TextRun({ text, font: "Arial", size: 22, bold: true, color: ACCENT_ORANGE })]
  });
}

function para(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 80, after: 80 },
    children: [new TextRun({ text, font: "Arial", size: 20, color: DARK, ...opts })]
  });
}

function bullet(text, bold = false) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text, font: "Arial", size: 20, color: DARK, bold })]
  });
}

function subbullet(text) {
  return new Paragraph({
    numbering: { reference: "subbullets", level: 0 },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, font: "Arial", size: 19, color: GRAY })]
  });
}

function tip(label, text) {
  return new Paragraph({
    spacing: { before: 80, after: 80 },
    children: [
      new TextRun({ text: `${label}: `, font: "Arial", size: 20, bold: true, color: ACCENT_ORANGE }),
      new TextRun({ text, font: "Arial", size: 20, color: DARK })
    ]
  });
}

function callout(text) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [new TableRow({
      children: [new TableCell({
        borders: { top: { style: BorderStyle.SINGLE, size: 8, color: BRAND_BLUE }, bottom: { style: BorderStyle.SINGLE, size: 8, color: BRAND_BLUE }, left: { style: BorderStyle.SINGLE, size: 20, color: BRAND_BLUE }, right: noBorder },
        shading: { fill: BRAND_LIGHT, type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 180, right: 120 },
        width: { size: 9360, type: WidthType.DXA },
        children: [new Paragraph({ children: [new TextRun({ text, font: "Arial", size: 20, color: BRAND_BLUE, italics: true })] })]
      })]
    })]
  });
}

function sectionDivider() {
  return new Paragraph({ spacing: { before: 240, after: 240 }, children: [] });
}

function headline(text) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [new TableRow({
      children: [new TableCell({
        borders: noBorders,
        shading: { fill: LIGHT_GRAY, type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 160, right: 160 },
        width: { size: 9360, type: WidthType.DXA },
        children: [new Paragraph({ children: [new TextRun({ text: `"${text}"`, font: "Arial", size: 20, color: DARK, italics: true })] })]
      })]
    })]
  });
}

function spacer() {
  return new Paragraph({ spacing: { before: 80, after: 80 }, children: [new TextRun("")] });
}

const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }]
      },
      {
        reference: "subbullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u25E6", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1080, hanging: 360 } } } }]
      },
      {
        reference: "numbering",
        levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }]
      }
    ]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 20 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 30, bold: true, font: "Arial", color: BRAND_BLUE }, paragraph: { spacing: { before: 480, after: 160 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 24, bold: true, font: "Arial", color: BRAND_BLUE }, paragraph: { spacing: { before: 320, after: 120 }, outlineLevel: 1 } },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
      }
    },
    children: [

      // === COVER ===
      new Paragraph({ spacing: { before: 480, after: 160 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "LinkedIn Growth & Visibility Strategy", font: "Arial", size: 40, bold: true, color: BRAND_BLUE })] }),
      new Paragraph({ spacing: { before: 0, after: 80 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "UTKARSH BHARDWAJ", font: "Arial", size: 28, bold: true, color: ACCENT_ORANGE })] }),
      new Paragraph({ spacing: { before: 0, after: 80 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Senior Angular + Full Stack Engineer  |  Enterprise Platform Engineering", font: "Arial", size: 22, italics: true, color: GRAY })] }),
      new Paragraph({ spacing: { before: 0, after: 320 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Targeting 35+ LPA  |  2026 India Tech Market", font: "Arial", size: 20, color: GRAY })] }),
      new Paragraph({ children: [new PageBreak()] }),

      // ============================================================
      // SECTION 1: PROFILE REVIEW STRATEGY
      // ============================================================
      h1("1. LinkedIn Profile Review Strategy"),
      para("Your profile is your 24/7 recruiter-facing landing page. For senior engineers targeting 35+ LPA in 2026, the bar is high — and most profiles fail at the basics. Here is how yours should look to maximize visibility, credibility, and callbacks."),
      spacer(),

      h2("What Recruiters Scan First (In Order)"),
      bullet("Profile photo — first 0.3 seconds of trust formation"),
      bullet("Headline — determines if they click or scroll past"),
      bullet("Current role + company — signals seniority level immediately"),
      bullet("About section — first 2–3 lines visible without expanding"),
      bullet("Experience titles + company names — scanned in under 10 seconds"),
      bullet("Skills section — directly feeds ATS keyword matching"),
      bullet("Featured section — differentiates you from 1000 other profiles"),
      spacer(),

      h2("Common Mistakes Senior Engineers Make"),
      bullet("Headline says only 'Senior Software Engineer at LTTS' — zero differentiation"),
      bullet("About section starts with 'I am a...' — weak, generic, skipped"),
      bullet("No numbers in experience bullets — sounds like job description, not impact"),
      bullet("Skills section is empty or has generic terms like 'Software Development'"),
      bullet("No featured content — missed credibility opportunity"),
      bullet("No Open to Work signal OR it is set incorrectly (visible to all, not just recruiters)"),
      bullet("90-day notice period buried or unstated — causes recruiter hesitation"),
      spacer(),

      h2("How to Look Strong Without Sounding Fake"),
      tip("Rule 1", "Use numbers you can defend in an interview — your resume already has great ones (~35% render improvement, ~69% build time reduction, ~95% HMR improvement). Use them everywhere."),
      tip("Rule 2", "Claim Angular specialist + backend collaborator, NOT backend architect. Your 70/30 split is a strength, not a gap."),
      tip("Rule 3", "Lead with outcomes, not activities. 'Reduced load time by 35%' beats 'Worked on performance optimisation.'"),
      tip("Rule 4", "Mention enterprise scale (global oil & gas clients, 100k+ nodes, production systems) — this signals senior-level context."),
      tip("Rule 5", "Keep tone professional but human. Avoid buzzword stacking without context. 'Micro Frontend architect' is stronger than 'experienced in modern frontend paradigms.'"),
      spacer(),

      callout("Your strongest differentiators: the measurable build time wins, Micro Frontend platform at LTTS, Angular v6→v20 upgrade with zero regressions, and AI productivity tools. Use these as anchors throughout your profile."),
      spacer(),

      // ============================================================
      // SECTION 2: HEADLINE OPTIONS
      // ============================================================
      new Paragraph({ children: [new PageBreak()] }),
      h1("2. Best LinkedIn Headline Options"),
      para("Your headline is the single highest-impact field on LinkedIn. It appears in search results, notifications, connection requests, and comments. Recruiters search by keywords in headlines. You have 220 characters — use all of them."),
      spacer(),

      h2("Top ATS + Recruiter-Optimized Headlines"),
      spacer(),

      h3("Option A — Angular Specialist Positioning (Recommended Primary)"),
      headline("Senior Angular Engineer | TypeScript · RxJS · Signals · Micro Frontends | Enterprise Platform Engineering | Full Stack (Java/Node.js) | 10 Yrs"),
      spacer(),

      h3("Option B — Full Stack Platform Engineering"),
      headline("Senior Full Stack Engineer | Angular v20 · React · TypeScript | Micro Frontend Architecture | Java Microservices | Enterprise UI Platform | LTTS/Oracle"),
      spacer(),

      h3("Option C — Keyword-Dense ATS Version"),
      headline("Angular Developer | TypeScript · RxJS · Signals · Standalone Components | Frontend Performance | Micro Frontends | Full Stack | 10+ Years Enterprise"),
      spacer(),

      h3("Option D — Platform Engineering Brand"),
      headline("UI Platform Engineer | Angular · Micro Frontends · Module Federation | Enterprise Workflow Systems | TypeScript · Java · REST APIs | 10 Yrs"),
      spacer(),

      h3("Option E — Concise Premium Feel"),
      headline("Senior Angular + Full Stack Engineer | Micro Frontends · Signals · RxJS | Enterprise Platform | 10 Years | Open to 35+ LPA Roles"),
      spacer(),

      h2("Headline Keyword Analysis"),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [3120, 2160, 4080],
        rows: [
          new TableRow({ children: [
            new TableCell({ borders, shading: { fill: BRAND_BLUE, type: ShadingType.CLEAR }, width: { size: 3120, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Keyword", font: "Arial", size: 20, bold: true, color: WHITE })] })] }),
            new TableCell({ borders, shading: { fill: BRAND_BLUE, type: ShadingType.CLEAR }, width: { size: 2160, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Search Volume", font: "Arial", size: 20, bold: true, color: WHITE })] })] }),
            new TableCell({ borders, shading: { fill: BRAND_BLUE, type: ShadingType.CLEAR }, width: { size: 4080, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Why It Matters", font: "Arial", size: 20, bold: true, color: WHITE })] })] }),
          ]}),
          ...[
            ["Angular Engineer", "Very High", "Primary recruiter search term for your role"],
            ["TypeScript", "High", "Almost always paired with Angular in recruiter filters"],
            ["RxJS", "High", "Differentiates Angular specialist from generic FE dev"],
            ["Micro Frontends", "High", "Hot in 2026, few candidates have real production exp"],
            ["Full Stack Engineer", "Very High", "Widens your reach to full stack JD searches"],
            ["Platform Engineering", "Medium-High", "Growing segment, maps to your LTTS work"],
            ["Signals", "Medium", "New Angular paradigm, shows you are current"],
            ["Module Federation", "Medium", "Technical depth signal for senior roles"],
            ["Enterprise", "High", "Filters you into high-salary enterprise searches"],
          ].map((row, i) => new TableRow({ children: row.map((cell, j) => new TableCell({
            borders,
            shading: { fill: i % 2 === 0 ? WHITE : LIGHT_GRAY, type: ShadingType.CLEAR },
            width: { size: [3120, 2160, 4080][j], type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [new Paragraph({ children: [new TextRun({ text: cell, font: "Arial", size: 19, color: DARK })] })]
          }))}))
        ]
      }),
      spacer(),

      // ============================================================
      // SECTION 3: ABOUT SECTION
      // ============================================================
      new Paragraph({ children: [new PageBreak()] }),
      h1("3. About Section Optimization"),
      para("The About section is your pitch. LinkedIn shows only the first 265 characters before the 'see more' click. Your opening line must hook the recruiter immediately."),
      spacer(),

      h2("Version A — Concise (Recommended for Most Recruiters)"),
      callout("Senior Angular + Full Stack Engineer with 10 years building enterprise-grade UI platforms at scale.\n\nI specialise in Angular (v4–v20), TypeScript, RxJS, and Signals-based architecture — delivering production systems for global clients in oil & gas, enterprise workflow, and data platforms.\n\nAt LTTS (Halliburton account), I:\n— Reduced Angular build time by ~69% and HMR rebuild time by ~95%\n— Improved UI rendering by ~35% on hierarchical tree datasets with 100k+ nodes\n— Designed and shipped a Micro Frontend platform using Webpack Module Federation (4+ independently deployable apps)\n— Led an Angular v6 → v20 upgrade with zero production regressions\n— Built internal AI productivity tools using LLM APIs to automate engineering workflows\n\nI contribute to backend teams on REST API contracts and Java microservice data flows — enough to work independently across the stack, but my home is the frontend platform.\n\n📍 Bengaluru | Open to Senior Angular, Full Stack, and Platform Engineering roles | 35+ LPA\n\n🔑 Angular · TypeScript · RxJS · Signals · Micro Frontends · Module Federation · Webpack · React · Java · Node.js · REST APIs · PostgreSQL · Redis · GitLab CI/CD · Performance Optimisation · Enterprise Platform Engineering"),
      spacer(),

      h2("Version B — Detailed Storytelling Version"),
      callout("I build enterprise frontend platforms that other engineers build on top of.\n\nWith 10 years in full stack engineering — 70% frontend-focused — I have spent most of my career solving the hard parts of large-scale Angular: performance on complex data trees, cross-team Micro Frontend architecture, build pipeline optimisation, and production debugging at scale.\n\nAt LTTS (Client: Halliburton), I lead frontend platform engineering for a global oil & gas PDM system used by KOC and ONGC. Some highlights:\n\n⚡ Performance: Cut Angular build time ~69%, HMR time ~95%, and UI render time ~35% on 100k+ node hierarchies — using OnPush change detection, virtual scrolling, and Webpack restructuring.\n\n🏗 Architecture: Designed a Micro Frontend platform with Webpack Module Federation — 4+ standalone Angular apps with shared runtime, improving release velocity across teams.\n\n🔄 Modernisation: Led Angular v6 → v20 migration with Signals, Standalone Components, and zero production regressions across a large enterprise codebase.\n\n🤖 AI: Built internal LLM-powered tools (Sprint Sense, AI Engineering Assistant) to automate sprint reporting and onboarding.\n\nOn the backend, I collaborate closely with Java/JAX-RS microservice teams on REST API contracts, async data flows, and PostgreSQL/JSONB data handling. I can hold my own across the stack.\n\nPreviously: 3 years at Oracle building enterprise Angular UI modules and backend REST APIs.\n\nI am open to Senior Angular, Full Stack, Enterprise Platform, or UI Platform Engineering roles in Bengaluru (or remote). Currently on a 90-day notice period.\n\n💬 Happy to connect with hiring managers, tech leads, and recruiters working on interesting platform challenges.\n\n#Angular #TypeScript #MicroFrontends #FullStack #PlatformEngineering"),
      spacer(),

      h2("ATS Keyword-Dense Version (for Copy-Pasting to Applications)"),
      para("Use this as your resume summary for job portals like Naukri, Foundit, LinkedIn Easy Apply:"),
      callout("Senior Angular + Full Stack Engineer | 10 years | Angular v4-v20 | TypeScript | RxJS | Signals | Standalone Components | Micro Frontends | Webpack Module Federation | Frontend Performance Optimization | Change Detection | Lazy Loading | Virtual Scrolling | React | Java | JAX-RS | Node.js | REST API | PostgreSQL | JSONB | Redis | Docker | GitLab CI/CD | Cypress | Playwright | Enterprise Platform Engineering | SPA Architecture | LLM Integration | Agile Scrum Master | Mentoring | Code Reviews"),
      spacer(),

      // ============================================================
      // SECTION 4: EXPERIENCE SECTION
      // ============================================================
      new Paragraph({ children: [new PageBreak()] }),
      h1("4. Experience Section Optimization"),
      spacer(),

      h2("LinkedIn Experience vs. Resume: Key Differences"),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [4680, 4680],
        rows: [
          new TableRow({ children: [
            new TableCell({ borders, shading: { fill: BRAND_BLUE, type: ShadingType.CLEAR }, width: { size: 4680, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Resume", font: "Arial", size: 20, bold: true, color: WHITE })] })] }),
            new TableCell({ borders, shading: { fill: BRAND_BLUE, type: ShadingType.CLEAR }, width: { size: 4680, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "LinkedIn", font: "Arial", size: 20, bold: true, color: WHITE })] })] }),
          ]}),
          ...[
            ["Concise, PDF-first", "Conversational, searchable, scannable"],
            ["1-2 pages max", "No length limit — use it"],
            ["Bullet-dense", "Short bullets + context sentences"],
            ["Keyword-dense", "Keyword-rich but readable"],
            ["No hashtags or tags", "Use hashtags and skill tags"],
            ["ATS-parsed as text", "Indexed by LinkedIn search algorithm"],
            ["No multimedia", "Add media, links, GitHub, writeups"],
          ].map((row, i) => new TableRow({ children: row.map((cell, j) => new TableCell({
            borders,
            shading: { fill: i % 2 === 0 ? WHITE : LIGHT_GRAY, type: ShadingType.CLEAR },
            width: { size: 4680, type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [new Paragraph({ children: [new TextRun({ text: cell, font: "Arial", size: 19, color: DARK })] })]
          }))}))
        ]
      }),
      spacer(),

      h2("LTTS (Feb 2021 – Present) — Recommended LinkedIn Format"),
      para("Title: Senior Full Stack Engineer & Project Lead | Angular Platform | Scrum Master", { bold: true }),
      spacer(),
      para("Opening context sentence (LinkedIn-specific addition):"),
      callout("Leading frontend platform engineering for a large-scale enterprise PDM system used by global oil & gas clients (KOC, ONGC) — end-to-end, from architecture through production support."),
      spacer(),
      para("Then use short, impactful bullets. Expand vs. resume guidance:"),
      bullet("EXPAND: Micro Frontend design decision — explain WHY you chose Module Federation, what problem it solved"),
      bullet("EXPAND: Build time wins — add context (this was blocking developer productivity)"),
      bullet("EXPAND: Angular upgrade — mention team size, codebase scale, migration strategy"),
      bullet("SHORTEN: Tool mentions (Cypress, Husky, ESLint) — list as a brief sentence, not separate bullets"),
      bullet("ADD: Skills tags at the bottom of each role — LinkedIn allows adding skills to each experience entry"),
      spacer(),

      h2("Oracle & Infosys — Recommended Approach"),
      bullet("Keep to 3–4 bullets maximum — recruiters care most about your current role"),
      bullet("Lead with the most impressive metric or outcome from each role"),
      bullet("Add skills tags (Angular, TypeScript, REST APIs, Java) to both roles for ATS coverage"),
      bullet("Infosys: Highlight the Cassandra → AWS migration — it is a strong data engineering signal"),
      spacer(),

      h2("Skills Section Strategy"),
      para("This is directly searched by recruiters. Prioritize in this order:"),
      bullet("Angular (top skill — pin it)", true),
      bullet("TypeScript (pin it)", true),
      bullet("RxJS (pin it)", true),
      bullet("Micro Frontends, Webpack Module Federation, Angular Performance"),
      bullet("React, Node.js, Java, REST APIs, Spring Boot (even if limited)"),
      bullet("PostgreSQL, Redis, Docker, GitLab CI/CD"),
      bullet("Agile, Scrum, Technical Leadership, Code Review, Mentoring"),
      para("Get endorsements for Angular, TypeScript, and RxJS from colleagues — endorsed skills rank higher in search."),
      spacer(),

      // ============================================================
      // SECTION 5: KEYWORD STRATEGY
      // ============================================================
      new Paragraph({ children: [new PageBreak()] }),
      h1("5. LinkedIn Keyword Strategy"),
      spacer(),

      h2("Master Keyword List by Category"),

      h3("Tier 1 — Must-Have (High Recruiter Search Volume)"),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [4680, 4680],
        rows: [
          new TableRow({ children: [
            new TableCell({ borders, shading: { fill: BRAND_BLUE, type: ShadingType.CLEAR }, width: { size: 4680, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Category", font: "Arial", size: 20, bold: true, color: WHITE })] })] }),
            new TableCell({ borders, shading: { fill: BRAND_BLUE, type: ShadingType.CLEAR }, width: { size: 4680, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Keywords", font: "Arial", size: 20, bold: true, color: WHITE })] })] }),
          ]}),
          ...[
            ["Angular", "Angular, Angular Developer, Angular Engineer, Angular 17, Angular 18, Angular 20, Angular Signals, Standalone Components, Angular Performance, Change Detection, OnPush"],
            ["Full Stack", "Full Stack Engineer, Full Stack Developer, Frontend Engineer, UI Engineer, Web Application Engineer"],
            ["TypeScript/RxJS", "TypeScript, RxJS, Observable, Reactive Programming, ES6, JavaScript"],
            ["Architecture", "Micro Frontend, Module Federation, Webpack, SPA Architecture, Frontend Architecture, Component Architecture"],
            ["Platform Engineering", "Platform Engineer, UI Platform, Frontend Platform, Enterprise Platform, Workflow Platform"],
            ["Backend (Credible)", "Java, JAX-RS, Node.js, REST API, Microservices, Spring Boot, Backend Development"],
            ["Data", "PostgreSQL, JSONB, Redis, MongoDB, Database Design, API Integration"],
            ["DevOps", "CI/CD, GitLab, Docker, Jenkins, Build Optimisation"],
            ["Testing", "Cypress, Playwright, E2E Testing, Unit Testing, Test Automation"],
            ["Soft Skills", "Technical Leadership, Scrum Master, Agile, Mentoring, Code Review, SDLC"],
          ].map((row, i) => new TableRow({ children: row.map((cell, j) => new TableCell({
            borders,
            shading: { fill: i % 2 === 0 ? WHITE : LIGHT_GRAY, type: ShadingType.CLEAR },
            width: { size: 4680, type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [new Paragraph({ children: [new TextRun({ text: cell, font: "Arial", size: 18, color: DARK })] })]
          }))}))
        ]
      }),
      spacer(),

      h2("Hidden High-Value Keywords (Less Common, High Signal)"),
      bullet("Angular Signals — very few candidates can legitimately claim this"),
      bullet("Webpack Module Federation — technical depth, not overused"),
      bullet("Virtual Scrolling — performance specialist signal"),
      bullet("LLM Integration / AI-Assisted Development — 2026 differentiator"),
      bullet("Sprint Facilitation — non-obvious PM-adjacent signal"),
      bullet("Zero Regression Migration — shows mature engineering mindset"),
      bullet("100k+ nodes / large data hierarchies — scale-of-work signal"),
      bullet("JSONB — specific PostgreSQL knowledge, signals backend credibility"),
      spacer(),

      h2("Keyword Placement Strategy"),
      bullet("Headline: 5–8 of your most searched keywords (Angular, TypeScript, Micro Frontends, Full Stack, Platform Engineering)"),
      bullet("About section: Dense keyword block at the end. Use a 🔑 emoji + comma-separated list"),
      bullet("Experience bullets: Natural keyword integration — 'Angular v20', 'RxJS-optimised', 'Webpack Module Federation'"),
      bullet("Skills section: All Tier 1 keywords, pinned top 3 = Angular, TypeScript, RxJS"),
      bullet("Certifications: Add Scrum Master, AI/ML keywords here"),
      bullet("Post content: Use hashtags #Angular #MicroFrontends #TypeScript #PlatformEngineering in every post"),
      spacer(),

      h2("How Recruiters Search in 2026"),
      callout("In 2026, Indian tech recruiters use LinkedIn Recruiter Lite or Premium. They filter by: Location (Bengaluru), Current Title keywords, Skills tags, Years of experience, Companies (Oracle/LTTS signal quality). The algorithm boosts profiles that: are active (posting/commenting), have high connection counts, have endorsed skills, and match keyword density. Your goal is to appear in the top 15 results for 'Angular Engineer Bengaluru'."),
      spacer(),

      // ============================================================
      // SECTION 6: BANNER
      // ============================================================
      new Paragraph({ children: [new PageBreak()] }),
      h1("6. LinkedIn Banner Strategy"),
      spacer(),
      para("Your banner is seen every time someone visits your profile. It signals your brand before they read a single word. At the senior level, a thoughtful banner separates you from 90% of engineers who leave it blank or use the default."),
      spacer(),

      h2("Recommended Banner Concept"),
      h3("Concept 1 — Minimalist Dark Tech (Best for Senior Engineering)"),
      bullet("Background: Deep navy (#0D1B3E) or dark charcoal (#1A1A2E)"),
      bullet("Left side: Your name in clean white + 'Senior Angular + Platform Engineer' in muted gray"),
      bullet("Center/Right: Angular logo + TypeScript + small icons (Micro Frontends, Performance)"),
      bullet("Bottom strip: Thin orange or blue accent line with 3–4 keywords: Angular  ·  Micro Frontends  ·  Enterprise Platform  ·  10 Years"),
      bullet("No stock photos, no gradients — clean, confident, technical"),
      spacer(),

      h3("Concept 2 — Architecture Blueprint Style"),
      bullet("Background: Dark blue with faint circuit/node pattern (very subtle)"),
      bullet("Left: Name + title in white"),
      bullet("Right: A simplified Micro Frontend architecture diagram (boxes + lines, 3–4 components)"),
      bullet("This immediately signals: this person thinks in systems"),
      spacer(),

      h3("Concept 3 — Metrics Banner (High Impact)"),
      bullet("Background: Clean dark tone"),
      bullet("Three stat blocks: '~69% Faster Builds' | '~35% Faster Rendering' | '10+ Years Enterprise'"),
      bullet("Subtitle: 'Senior Angular + Full Stack Engineer | Enterprise Platform'"),
      bullet("Works well because it shows impact before they read anything"),
      spacer(),

      h2("Banner Design Specifications"),
      bullet("Size: 1584 × 396 pixels (LinkedIn standard)"),
      bullet("Tools: Canva (free, has LinkedIn banner templates), Figma, Adobe Express"),
      bullet("Colors: Navy (#1B4F9B), White (#FFFFFF), Orange accent (#E07B30)"),
      bullet("Font: Inter, Roboto, or DM Sans — modern, clean, tech-appropriate"),
      bullet("Avoid: photos of yourself, stock images of laptops/handshakes, too much text"),
      bullet("Include Angular logo? Yes, it is well-recognized and brand-safe for personal branding"),
      spacer(),

      callout("Recommendation: Start with Concept 1. Use Canva, choose a dark LinkedIn banner template, add your name/title on the left, Angular + MF icons on the right, and a keyword strip at the bottom. Takes 20 minutes and immediately elevates your profile."),
      spacer(),

      // ============================================================
      // SECTION 7: PROFILE PHOTO
      // ============================================================
      h1("7. Profile Photo Strategy"),
      spacer(),

      h2("What an Ideal Senior Engineer Photo Looks Like"),
      bullet("Framing: Head and shoulders (face fills 60–70% of the frame)"),
      bullet("Background: Plain white, light gray, or very slightly blurred neutral office/outdoor setting. Avoid busy backgrounds."),
      bullet("Expression: Confident, approachable smile — not stiff corporate. Think 'I am good at what I do and easy to work with.'"),
      bullet("Eye contact: Direct eye contact with camera"),
      bullet("Lighting: Natural light from a window is ideal. Front-lit, no harsh shadows."),
      spacer(),

      h2("Dress Code for Senior Tech Roles in India 2026"),
      bullet("Smart casual: clean collared shirt (solid color — navy, white, gray, dark green), no tie needed"),
      bullet("Avoid: T-shirts with logos/graphics, formal suit unless interviewing at a bank, heavily patterned clothing"),
      bullet("Colors that read well on LinkedIn: Navy, white, dark gray, forest green — contrast well against light backgrounds"),
      spacer(),

      h2("Technical Tips"),
      bullet("Resolution: At least 400 × 400 px (LinkedIn crops to a circle at 400px)"),
      bullet("File size: Under 8MB"),
      bullet("Avoid: Sunglasses, hats, group photos, selfies with visible arm, heavily filtered photos"),
      bullet("Test: Does your face look confident and recognizable as a thumbnail (50x50px)? That is what recruiters see in search results."),
      spacer(),

      // ============================================================
      // SECTION 8: FEATURED SECTION
      // ============================================================
      new Paragraph({ children: [new PageBreak()] }),
      h1("8. Featured Section Strategy"),
      para("The Featured section is seen by 100% of profile visitors who scroll. For a senior engineer targeting premium roles, this is your portfolio showcase. Most engineers leave it empty — do not."),
      spacer(),

      h2("Priority Featured Items (Ranked)"),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [360, 3240, 5760],
        rows: [
          new TableRow({ children: [
            new TableCell({ borders, shading: { fill: BRAND_BLUE, type: ShadingType.CLEAR }, width: { size: 360, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "#", font: "Arial", size: 20, bold: true, color: WHITE })] })] }),
            new TableCell({ borders, shading: { fill: BRAND_BLUE, type: ShadingType.CLEAR }, width: { size: 3240, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Item", font: "Arial", size: 20, bold: true, color: WHITE })] })] }),
            new TableCell({ borders, shading: { fill: BRAND_BLUE, type: ShadingType.CLEAR }, width: { size: 5760, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Why / What to Include", font: "Arial", size: 20, bold: true, color: WHITE })] })] }),
          ]}),
          ...[
            ["1", "Micro Frontend Architecture Post or Article", "Write a LinkedIn article or post about your Module Federation design. Screenshot the architecture diagram. This is your highest-credibility content — few people have shipped this in production."],
            ["2", "Angular Performance Case Study", "Post or PDF: '3 techniques that cut our Angular render time by 35%.' Cover OnPush, virtual scrolling, and RxJS optimisation. No code needed — architecture explanation is enough."],
            ["3", "GitHub Profile or Repository", "Link to a public GitHub repo — even a demo/sample project showing Angular Signals, Standalone Components, or a Micro Frontend setup. Signals production-level thinking."],
            ["4", "Angular v6 → v20 Migration Write-up", "A short post or article: 'What we learned migrating Angular v6 to v20 in production.' Include zero-regression approach, Signals adoption, and team impact. Extremely rare content."],
            ["5", "AI Engineering Tools Post", "Short post about Sprint Sense or AI Engineering Assistant. How you built it, what LLM/API you used, what problem it solved. Shows you are working with AI in 2026 — a differentiator."],
            ["6", "Resume PDF (Optional)", "Pin your resume as a PDF so recruiters can download it directly. Add a professional cover image for the thumbnail."],
          ].map((row, i) => new TableRow({ children: row.map((cell, j) => new TableCell({
            borders,
            shading: { fill: i % 2 === 0 ? WHITE : LIGHT_GRAY, type: ShadingType.CLEAR },
            width: { size: [360, 3240, 5760][j], type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [new Paragraph({ children: [new TextRun({ text: cell, font: "Arial", size: 18, color: DARK })] })]
          }))}))
        ]
      }),
      spacer(),

      // ============================================================
      // SECTION 9: CONTENT STRATEGY
      // ============================================================
      new Paragraph({ children: [new PageBreak()] }),
      h1("9. Daily/Weekly Content Strategy — 90 Days"),
      para("Consistent, relevant content is the fastest way to build recruiter visibility without spending money. You do not need to go viral — you need to be consistently visible to the right 500 people (recruiters, hiring managers, senior engineers in target companies)."),
      spacer(),

      h2("Best Post Formats in 2026"),
      bullet("Short insight posts (3–7 lines): Highest reach, lowest effort. Best for opinions, quick lessons."),
      bullet("Carousel posts (5–10 slides): High saves and shares, strong for technical topics. Use Canva."),
      bullet("Storytelling posts: 'Here is what happened when...' structure. High engagement from non-technical readers."),
      bullet("Debugging posts: 'Spent 3 hours on this bug. Here is what I found.' Extremely relatable."),
      bullet("Architecture explanation posts: 1 concept, 1 diagram, 3 bullet insights. Builds authority."),
      bullet("Before/After posts: 'Before: build took 3 mins. After: 52 seconds. Here is how.' Very high saves."),
      spacer(),

      h2("90-Day Post Calendar"),

      h3("Month 1 — Foundation (Weeks 1–4)"),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [1440, 2880, 5040],
        rows: [
          new TableRow({ children: [
            new TableCell({ borders, shading: { fill: BRAND_BLUE, type: ShadingType.CLEAR }, width: { size: 1440, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Week", font: "Arial", size: 20, bold: true, color: WHITE })] })] }),
            new TableCell({ borders, shading: { fill: BRAND_BLUE, type: ShadingType.CLEAR }, width: { size: 2880, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Format", font: "Arial", size: 20, bold: true, color: WHITE })] })] }),
            new TableCell({ borders, shading: { fill: BRAND_BLUE, type: ShadingType.CLEAR }, width: { size: 5040, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Topic", font: "Arial", size: 20, bold: true, color: WHITE })] })] }),
          ]}),
          ...[
            ["Week 1", "Short insight post", "The one Angular change that halved our render time (OnPush + virtual scrolling)"],
            ["Week 1", "Story post", "'I spent 10 years in Angular. Here is what nobody tells you about scaling it.' Personal hook."],
            ["Week 2", "Carousel", "5 slides: 'Angular Signals vs RxJS — which one and when?' Practical, searchable."],
            ["Week 2", "Debugging story", "'3-hour debug session. Here is the Angular change detection bug I finally cracked.'"],
            ["Week 3", "Architecture post", "Micro Frontend diagram + 3-line explanation. 'Why we split our Angular app into 4 deployable units.'"],
            ["Week 3", "Short post", "Opinion: 'Standalone Components in Angular are the future. Here is what changed in production for us.'"],
            ["Week 4", "Before/after post", "'Angular build: 4 min 20s → 1 min 22s. Here is what we changed.' With screenshot if possible."],
            ["Week 4", "Career/soft skills", "'What 3 years as a Scrum Master taught me about building better software.' Humanizes you."],
          ].map((row, i) => new TableRow({ children: row.map((cell, j) => new TableCell({
            borders,
            shading: { fill: i % 2 === 0 ? WHITE : LIGHT_GRAY, type: ShadingType.CLEAR },
            width: { size: [1440, 2880, 5040][j], type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [new Paragraph({ children: [new TextRun({ text: cell, font: "Arial", size: 18, color: DARK })] })]
          }))}))
        ]
      }),
      spacer(),

      h3("Month 2 — Authority Building (Weeks 5–8)"),
      bullet("Week 5: 'How we integrated React inside an Angular Micro Frontend shell — and why it worked'"),
      bullet("Week 5: Carousel: '5 RxJS patterns I use every day in Angular (and 2 I stopped using)'"),
      bullet("Week 6: 'The Angular v6 → v20 migration story — what broke, what we fixed, what we learned'"),
      bullet("Week 6: Short post on Angular Signals + before/after reactivity model"),
      bullet("Week 7: 'How I built an AI sprint assistant for our engineering team using LLM APIs'"),
      bullet("Week 7: Architecture explainer: JSONB in PostgreSQL — when to use it vs normalized tables"),
      bullet("Week 8: Debugging story: 'A race condition in Angular + RxJS that took us 2 days to find'"),
      bullet("Week 8: Opinion post: 'Why frontend engineers should understand Redis caching (even if they never configure it)'"),
      spacer(),

      h3("Month 3 — Credibility + Reach (Weeks 9–12)"),
      bullet("Week 9: 'How to set up Cypress for a large Angular app — what we got wrong the first time'"),
      bullet("Week 9: Career story: '10 years as an engineer — the 3 decisions that shaped my trajectory'"),
      bullet("Week 10: Technical deep-dive carousel: 'Webpack Module Federation in 5 slides — the mental model'"),
      bullet("Week 10: Opinion: 'AI-assisted development in 2026 — what's hype vs what genuinely saves time'"),
      bullet("Week 11: Share a debugging approach or mental model for performance profiling in Angular DevTools"),
      bullet("Week 11: 'How I mentor junior engineers through code reviews — the 3 things I always look for'"),
      bullet("Week 12: Reflection post: '90 days of posting on LinkedIn — what I learned' (meta-post, high reach)"),
      bullet("Week 12: Your Micro Frontend architecture diagram with a 3-para explanation"),
      spacer(),

      // ============================================================
      // SECTION 10: VIRAL/RECRUITER REACH
      // ============================================================
      new Paragraph({ children: [new PageBreak()] }),
      h1("10. Viral & Recruiter Reach Strategy"),
      spacer(),

      h2("How the LinkedIn Algorithm Works in 2026"),
      para("LinkedIn uses an engagement velocity model — posts that get likes/comments within the first 60–90 minutes are shown to more people. Key signals:"),
      bullet("Comments > Likes > Reposts for algorithm boost"),
      bullet("Dwell time matters — long-form posts that people read fully perform better"),
      bullet("Profile strength (All-Star status) boosts search visibility independently of content"),
      bullet("First-degree connections who engage carry more weight than strangers"),
      bullet("Native content (articles, carousels, text posts) outperforms external links"),
      spacer(),

      h2("Best Posting Times for India (2026)"),
      bullet("Tuesday–Thursday: Best days overall"),
      bullet("7:30–9:00 AM IST: Morning commute / pre-work peak"),
      bullet("12:00–1:30 PM IST: Lunch scroll window"),
      bullet("6:00–8:00 PM IST: Post-work peak — highest engagement for tech content"),
      bullet("Avoid: Friday evenings, weekends, Monday mornings"),
      spacer(),

      h2("Comment Strategy"),
      para("Commenting on others' posts is the most underrated growth lever. Target:"),
      bullet("Posts by Angular team members, Angular influencers, ThoughtWorks leaders, tech hiring managers at your target companies"),
      bullet("Comment with a substantive 2–3 line technical insight, not just 'Great post!'"),
      bullet("Your comment inherits the reach of the original post — every person who sees it may visit your profile"),
      bullet("Goal: 5 meaningful comments per day on relevant posts in your first 30 days"),
      spacer(),

      h2("Recruiter Outreach Strategy"),
      para("Proactive outreach is 3–5x more effective than waiting for inbound. Template approach:"),
      callout("Hi [Name], I noticed you recruit for senior frontend/full stack roles at [Company]. I'm a Senior Angular + Full Stack Engineer with 10 years of experience — most recently at LTTS building enterprise platform systems (Micro Frontends, Angular v6→v20 migration, ~69% build optimisation). I'm exploring 35+ LPA roles in Bengaluru. Happy to share my resume if the timing works for your team."),
      spacer(),
      bullet("Target: 5–10 recruiter outreaches per week"),
      bullet("Personalize each message — mention the company or a specific JD you saw"),
      bullet("Connect first, then message (higher acceptance than InMail cold messages)"),
      bullet("Follow recruiters at Naukri, TeamLease, Michael Page, Randstad, and company-internal HR at Flipkart, Razorpay, Atlassian, Thoughtworks, Publicis Sapient, Capgemini Engineering"),
      spacer(),

      h2("Connection Strategy"),
      bullet("Grow to 500+ connections as fast as possible — 500+ signals credibility and boosts search rank"),
      bullet("Connect with: ex-colleagues, LinkedIn classmates, Angular community members, speakers at Angular India events, people who engage with your posts"),
      bullet("Send personalized notes with connection requests — 2 lines about why you are connecting"),
      bullet("Avoid random mass-connection campaigns — LinkedIn's algorithm penalizes high rejection rates"),
      spacer(),

      // ============================================================
      // SECTION 11: JOB SEARCH OPTIMIZATION
      // ============================================================
      h1("11. Job Search Optimization"),
      spacer(),

      h2("Open to Work Settings"),
      bullet("Turn ON 'Open to Work' but set visibility to 'Recruiters only' (not 'All LinkedIn members')"),
      bullet("This hides the green frame from your profile photo (avoids signaling desperation to current employer) while making you visible in LinkedIn Recruiter searches"),
      bullet("Set preferred job titles to: Senior Angular Engineer, Senior Full Stack Engineer, UI Platform Engineer, Enterprise Platform Engineer, Frontend Architect"),
      bullet("Set preferred locations: Bengaluru (on-site), and select 'Remote' as well to maximize reach"),
      bullet("Set start date: In 3 months (signals notice period without being explicit)"),
      spacer(),

      h2("How to Handle 90-Day Notice Period"),
      callout("Do NOT hide your notice period. Include it clearly in your About section and in recruiter conversations. Many companies at 35+ LPA are fine with 90 days — they plan ahead. The risk is starting a process only to have it fall apart when the notice period surfaces. Being upfront saves everyone time and builds trust."),
      spacer(),
      para("Suggested wording for About section: 'Currently on a 90-day notice period. Available from [Month + Year].'"),
      spacer(),

      h2("Naukri.com Profile (Critical for India Market)"),
      bullet("Mirror your LinkedIn headline and About section on Naukri"),
      bullet("Update your resume on Naukri every 7–10 days to stay in 'Freshly Updated' searches"),
      bullet("Set profile visibility to maximum, enable recruiter calls"),
      bullet("Add all skills as searchable tags"),
      bullet("Set expected CTC clearly: 35–45 LPA"),
      spacer(),

      h2("Improving Callback Rate"),
      bullet("Ensure your resume file is ATS-parseable: simple formatting, no text boxes, no headers/footers with critical info"),
      bullet("Apply within 24 hours of a JD posting — early applications get more visibility"),
      bullet("Tailor your resume summary for each application — match their exact JD language"),
      bullet("For roles at Bengaluru product companies (Razorpay, Atlassian, Swiggy, Flipkart), reach out to the hiring manager on LinkedIn simultaneously with applying"),
      spacer(),

      // ============================================================
      // SECTION 12: CONTENT POSITIONING
      // ============================================================
      new Paragraph({ children: [new PageBreak()] }),
      h1("12. Content Positioning"),
      spacer(),

      h2("Topics to Become Known For"),
      bullet("Angular performance optimization (OnPush, virtual scrolling, lazy loading, build optimization) — your strongest proof points"),
      bullet("Micro Frontend architecture — specifically Module Federation. This is rare and marketable."),
      bullet("Angular upgrade strategies — v-to-v migrations at enterprise scale"),
      bullet("Angular Signals and modern Angular patterns — positions you as current and forward-looking"),
      bullet("Full stack API integration — REST contract design, frontend-backend data flows"),
      bullet("Engineering productivity and AI tools — differentiates you in 2026"),
      spacer(),

      h2("Topics to Avoid Overclaiming"),
      bullet("Kubernetes architecture or DevOps system design — you have familiarity, not ownership. Mention it as context, not expertise."),
      bullet("Backend system design at scale — your backend contributions are collaborative, not architect-level"),
      bullet("Distributed systems or high-availability infrastructure design"),
      bullet("AI/ML model training or data science — your AI work is LLM integration (which is legitimate and valuable, but different)"),
      spacer(),

      h2("How to Build Credibility Gradually"),
      para("The senior engineer credibility arc:"),
      bullet("Month 1: Share what you have already done — real numbers, real outcomes, no exaggeration"),
      bullet("Month 2: Teach from experience — explain how you solved problems others face"),
      bullet("Month 3: Have opinions — take a stance on Angular Signals vs RxJS, on Micro Frontends vs monolith"),
      bullet("Month 4+: Engage with the community — respond to questions, comment on Angular releases, share thoughts on industry trends"),
      spacer(),

      callout("The most credible signal: specificity. 'Reduced build time by ~69%' is 10x more credible than 'significantly improved build performance.' Use your real numbers. They are excellent."),
      spacer(),

      // ============================================================
      // SECTION 13: 30-DAY PLAN
      // ============================================================
      h1("13. 30-Day LinkedIn Growth Plan"),
      spacer(),

      h2("Week 1 — Foundation (Days 1–7)"),
      h3("Profile Optimization (Days 1–3)"),
      new Paragraph({ numbering: { reference: "numbering", level: 0 }, spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "Day 1: Update headline to Option A. Update profile photo (take new one if needed). Turn on Open to Work (recruiters only).", font: "Arial", size: 20, color: DARK })] }),
      new Paragraph({ numbering: { reference: "numbering", level: 0 }, spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "Day 2: Write and publish your About section (Version A from Section 3). Update your LTTS experience with LinkedIn-optimized format.", font: "Arial", size: 20, color: DARK })] }),
      new Paragraph({ numbering: { reference: "numbering", level: 0 }, spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "Day 3: Add all skills from Tier 1 keyword list. Request endorsements from 3–4 ex-colleagues for Angular, TypeScript, RxJS.", font: "Arial", size: 20, color: DARK })] }),
      spacer(),

      h3("First Content (Days 4–7)"),
      new Paragraph({ numbering: { reference: "numbering", level: 0 }, spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "Day 4: Post #1 — 'The one Angular change that halved our render time.' 5–8 lines, no code, hashtags: #Angular #PerformanceOptimization #Frontend", font: "Arial", size: 20, color: DARK })] }),
      new Paragraph({ numbering: { reference: "numbering", level: 0 }, spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "Day 5: Comment on 5 Angular/tech posts by influencers or hiring managers. Add 20 connection requests to ex-colleagues and Angular community members.", font: "Arial", size: 20, color: DARK })] }),
      new Paragraph({ numbering: { reference: "numbering", level: 0 }, spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "Day 6: Design LinkedIn banner in Canva (Concept 1). Upload it. Screenshot your updated profile and share it as a post ('Just refreshed my LinkedIn — open to senior Angular/platform engineering roles in Bengaluru').", font: "Arial", size: 20, color: DARK })] }),
      new Paragraph({ numbering: { reference: "numbering", level: 0 }, spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "Day 7: Send 5 recruiter outreach messages (use template from Section 10). Target: recruiters at staffing firms and mid-to-large tech companies.", font: "Arial", size: 20, color: DARK })] }),
      spacer(),

      h2("Week 2 — Momentum (Days 8–14)"),
      bullet("Day 8: Prepare Micro Frontend architecture diagram (pen/paper or draw.io). This will be your Featured section anchor."),
      bullet("Day 9: Post #2 — Storytelling post: 'I spent 10 years in Angular. Here is what nobody tells you about scaling it.'"),
      bullet("Day 10: Comment strategy: 5 meaningful comments. Connect with 20 more people."),
      bullet("Day 11: Start writing your Featured content: Angular performance case study (even a 300-word LinkedIn article counts)"),
      bullet("Day 12: Post #3 — Micro Frontend architecture diagram + explanation. Tag it #MicroFrontends #ModuleFederation #Angular"),
      bullet("Day 13: Send 5 more recruiter outreaches. Follow up on Week 1 outreaches if no response."),
      bullet("Day 14: Update Naukri/Foundit profile with same headline and about section. Enable recruiter contact."),
      spacer(),

      h2("Week 3 — Authority (Days 15–21)"),
      bullet("Day 15: Post #4 — Before/after build time post. Screenshot or specific numbers work best here."),
      bullet("Day 16: Comment + connect: 5 comments, 20 connections. Start following Angular core team members, Vikram Subramanian, Minko Gechev."),
      bullet("Day 17: Publish Featured item 1: Share the Micro Frontend architecture diagram as a standalone post, then pin it to Featured."),
      bullet("Day 18: Post #5 — Carousel: 'Angular Signals vs RxJS — which one and when?' 5 slides on Canva."),
      bullet("Day 19: Engage with all comments on your previous posts. Reply to every comment — this boosts algorithm reach."),
      bullet("Day 20: Send 5 more recruiter outreaches. You should have 15+ outreaches sent by now."),
      bullet("Day 21: Review your profile as a recruiter would — Google your name, look at your public profile, check if keywords appear in the snippet."),
      spacer(),

      h2("Week 4 — Conversion (Days 22–30)"),
      bullet("Day 22: Post #6 — Angular v6 → v20 migration story. 8–10 lines, focus on the human/engineering challenge."),
      bullet("Day 23: Connect with 20 senior engineers at your target companies (Flipkart, Razorpay, Atlassian, Thoughtworks, Publicis Sapient). Not for jobs — for community."),
      bullet("Day 24: Publish Featured item 2: Angular Performance case study article or detailed post."),
      bullet("Day 25: Post #7 — AI productivity tools: how you built Sprint Sense and why it mattered."),
      bullet("Day 26: Follow up on all recruiter outreaches. Many reply in Week 3–4."),
      bullet("Day 27: Post #8 — Debugging story. Real incident, real learning, real impact."),
      bullet("Day 28: Request 2–3 LinkedIn recommendations from LTTS manager, Oracle colleagues, or team members. Recommendations are a strong ATS and trust signal."),
      bullet("Day 29: Review analytics on all posts. Double down on the format/topic that got most engagement."),
      bullet("Day 30: Reflection post — '30 days, 8 posts, here is what I have learned about building in public as an engineer.' Vulnerable, relatable, high reach."),
      spacer(),

      h2("Weekly Recurring Actions (Ongoing After Day 30)"),
      bullet("Post 2x per week (Tuesday and Thursday recommended)"),
      bullet("Comment on 5 relevant posts per day (10 minutes daily)"),
      bullet("Send 5–7 recruiter/connection outreaches per week"),
      bullet("Refresh Naukri resume every 7–10 days"),
      bullet("Check profile analytics weekly — track search appearances and post reach"),
      spacer(),

      callout("The 30-day goal is not virality. It is to build a credible, keyword-optimized, active profile that ranks in recruiter searches and makes a strong impression when they visit. 8 quality posts, 500+ connections, 15+ recruiter outreaches, and a fully optimized profile will put you significantly ahead of most candidates at your level."),
      spacer(),

      // === CLOSING ===
      new Paragraph({ children: [new PageBreak()] }),
      new Paragraph({ spacing: { before: 480, after: 240 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Summary: Your Strongest Differentiation Points", font: "Arial", size: 28, bold: true, color: BRAND_BLUE })] }),
      spacer(),
      callout("Based on your resume, these are your top 5 credibility anchors. Lead with these everywhere — headline, about, posts, recruiter messages:\n\n1. ~69% Angular build time reduction + ~95% HMR reduction — exceptional, specific, verifiable\n2. Micro Frontend platform with 4+ independently deployable Angular apps — rare in India\n3. Angular v6 → v20 migration with zero production regressions — enterprise maturity signal\n4. ~35% rendering improvement on 100k+ node hierarchies — advanced Angular internals knowledge\n5. AI productivity tools (LLM integration) built for an engineering team — highly relevant in 2026\n\nYour positioning: Senior Angular specialist who thinks in systems, ships at enterprise scale, and understands the full stack enough to own end-to-end delivery. That is a rare and valuable combination at 35+ LPA."),
      spacer(),
      new Paragraph({ spacing: { before: 160, after: 160 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Good luck, Utkarsh. The profile and the numbers are strong. Execute the plan.", font: "Arial", size: 20, italics: true, color: GRAY })] }),

    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("linkedin_strategy_utkarsh.docx", buf);
  console.log("Done");
});