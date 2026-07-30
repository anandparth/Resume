// Resume generator. Setup: `npm install docx` (in this folder). Run: `node build_resume.js`
// Outputs Parth_Anand_Product_Designer_Resume.docx - after regenerating, also refresh the
// PDF (LibreOffice: `soffice --headless --convert-to pdf Parth_Anand_Product_Designer_Resume.docx`)
// and copy it over Parth_Anand_Resume_Product_Design.pdf, which is the file GitHub Pages
// actually serves and that parthanand.in links to.
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, LevelFormat,
  ExternalHyperlink, BorderStyle, TabStopType
} = require("docx");

const NAVY = "1A1A1A";
const RULE = "1A1A1A";
const FONT = "Calibri";
const LINE = { line: 242, lineRule: "auto" }; // ~1.01x, loosened after merging/trimming bullets in the final audit pass
const PAGE_MARGIN_L = 720, PAGE_MARGIN_R = 720, PAGE_WIDTH = 12240;
const RIGHT_TAB_POS = PAGE_WIDTH - PAGE_MARGIN_L - PAGE_MARGIN_R; // true usable-width right edge, NOT TabStopPosition.MAX (that constant is A4-sized and was landing ~1.2in short on this US Letter layout)

const doc = new Document({
  styles: {
    default: { document: { run: { font: FONT, size: 18 } } }, // 9pt
  },
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 260, hanging: 180 } } }
        }]
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 320, right: 720, bottom: 320, left: 720 } // 0.22in top/bottom, 0.5in sides
      }
    },
    children: [
      // NAME
      new Paragraph({
        spacing: { after: 20 },
        children: [ new TextRun({ text: "PARTH ANAND", bold: true, size: 32, font: FONT }) ]
      }),
      // TITLE
      new Paragraph({
        spacing: { after: 40 },
        children: [ new TextRun({ text: "Product Designer, UI/UX", size: 20, color: "404040" }) ]
      }),
      // CONTACT LINE
      new Paragraph({
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: RULE, space: 3 } },
        spacing: { after: 100 },
        children: [
          new TextRun({ text: "Noida, India  |  +91 72096 84084  |  ", size: 17 }),
          new ExternalHyperlink({ link: "mailto:parthanand1705@gmail.com", children: [new TextRun({ text: "parthanand1705@gmail.com", size: 17, style: "Hyperlink" })] }),
          new TextRun({ text: "  |  ", size: 17 }),
          new ExternalHyperlink({ link: "https://linkedin.com/in/parthanand21", children: [new TextRun({ text: "linkedin.com/in/parthanand21", size: 17, style: "Hyperlink" })] }),
          new TextRun({ text: "  |  ", size: 17 }),
          new ExternalHyperlink({ link: "https://parthanand.in", children: [new TextRun({ text: "parthanand.in", size: 17, style: "Hyperlink" })] }),
        ]
      }),

      sectionHeading("PROFESSIONAL SUMMARY"),
      new Paragraph({
        spacing: { after: 60, ...LINE },
        children: [ new TextRun({
          text: "Product designer with 5+ years designing workflow-heavy products for operational users: recruiter dashboards, fraud-prevention systems, and multi-step DeFi flows. Starts from the problem, not the solution, validating assumptions through research before committing to design. Shipped a high-volume recruiter dashboard for a hiring platform with 1M+ active job seekers, and led onboarding for a self-custody wallet with 1M+ users. Uses AI (Claude, Midjourney) daily to accelerate prototyping, assets, and research synthesis.",
          size: 18
        }) ]
      }),

      sectionHeading("CORE SKILLS"),
      new Paragraph({
        spacing: { after: 30 },
        children: [ new TextRun({
          text: "Product & UI/UX Design  |  Design Systems (Figma)  |  UX Research  |  Interaction Design  |  Prototyping  |  Usability Testing  |  Accessibility (WCAG)",
          size: 18, bold: true
        }) ]
      }),

      sectionHeading("EXPERIENCE"),

      ...roleBlock(
        "InfoEdge (Naukri.com) - Creative Designer, Product UI/UX",
        "Nov 2023 - Present",
        "Job Hai | Blue-collar hiring platform | 1M+ active job seekers",
        [
          "Designed four core screens of the Job Hai recruiter dashboard (candidate cards, candidate management, advanced filters, bulk hiring), lifting average actions per recruiter ~11% and overall platform actions ~8%.",
          "Built and scaled a Figma design system (components, variables, tokens, auto-layout) across 4+ product surfaces with PMs and engineers, standardizing consistency across agile sprint cycles.",
          "Mentored a junior designer on the Job Hai team, reviewing interaction-design work and onboarding them onto the team's Figma design system and workflow.",
          "Led a fraud-prevention initiative end to end: mapped 249 blacklisted recruiters and 653 fake listings across Delhi NCR, then modeled and presented a roadmap of 6 prioritized recommendations to senior stakeholders, projecting a 68% fraud reduction and a +22-point NPS lift; approved for implementation.",
          "Designed hiring workflows (job discovery, trust signals, fraud reporting) for a Hindi-first, low-digital-literacy user base, applying WCAG accessibility principles to lift task success and retention.",
          "Ran 12+ usability sessions combining moderated contextual inquiry, unmoderated testing, and Attention Insight eye-tracking to pressure-test design assumptions before rollout."
        ]
      ),

      ...roleBlock(
        "CoinDCX - Senior Executive Designer, Product Design",
        "Jun 2022 - Oct 2023",
        "Okto | Self-custody Web3 DeFi wallet | 1M+ users | iOS & Android",
        [
          "De-risked 3 critical flows ahead of build (empathy mapping, journey mapping, card sorting, heuristic evaluation), then designed multi-step swap, bridging, and earning flows covering edge cases, error states, and latency UI.",
          "Owned research-to-UI for Okto's first-time user onboarding, a core surface of a self-custody wallet serving 1M+ users across iOS and Android.",
          "Established a cross-platform Figma component library with design tokens to keep Okto's iOS and Android UI consistent as new flows shipped.",
          "Produced Lottie motion assets and visual content for two major Web3 community events (Unfold '22, '23)."
        ]
      ),

      ...roleBlock(
        "Independent - Motion Designer (Freelance)",
        "Feb 2022 - May 2022",
        null,
        [
          "Delivered end-to-end animation and motion-graphics packages for YouTube creators on freelance timelines, deepening expertise in After Effects and Lottie while applying motion design to scalable UI experiences and prototyping 3D interactions with Meta AR/VR technology."
        ]
      ),

      ...roleBlock(
        "HMLC (Harsh Mann Luxury Consultancy) - Lead Designer, UI/UX & Motion",
        "Aug 2021 - Feb 2022",
        "Boutique luxury retail & lifestyle consultancy",
        [
          "Designed responsive, mobile-first interfaces for 3+ retail clients, validating layout via A/B testing and Hotjar heatmap analysis.",
          "Achieved as much as a 22% increase in page conversion rate across tested variants.",
          "Owned the client-facing review process and developer handoff across those engagements - the primary point of contact translating client decisions into shippable specs for engineering."
        ]
      ),

      ...roleBlock(
        "Canon India - UI/UX Designer (Apprenticeship)",
        "Jan 2021 - Jul 2021",
        "Consumer camera redesign (concept project)",
        [
          "Redesigned Canon India's camera shopping experience, simplifying discovery, comparison, and purchase journeys across 5+ campaign pages."
        ],
        true
      ),

      sectionHeading("SKILLS & TOOLS"),
      skillLine("Design & Research:", "Wireframing, information architecture, motion and visual design, empathy mapping, card sorting, heuristic evaluation, A/B testing, affinity mapping, WCAG accessibility standards"),
      skillLine("Strategy & Collaboration:", "UX roadmapping, stakeholder presentations, design reviews, cross-functional alignment with PMs and engineering, client-facing communication"),
      skillLine("Systems & Frontend:", "Figma design systems (components, variables, tokens, auto-layout), HTML/CSS/JS, Three.js, WebGL, GSAP, Blender, Vercel deployment"),
      skillLine("Tools & AI Workflows:", "Framer, Lottie, Adobe Creative Suite, Miro, Hotjar, Amplitude, ProtoPie, Dovetail; AI-directed prototyping, asset production, and research synthesis (Claude, Midjourney)"),
      skillLine("Personal Project:", "papercade - solo-built, MIT-licensed open-source UI library (12 components, pixel x sketch aesthetic); strict TypeScript, automated CI/CD (typecheck, tests, auto-deploy), published to npm with a live demo - github.com/anandparth/papercade", true),

      sectionHeading("EDUCATION & CERTIFICATIONS"),
      ...eduItem("B.Des. in Fashion Communication - National Institute of Fashion Technology (NIFT), New Delhi", "2017 - 2021"),
      ...eduItem("Professional UX Design - University of Cambridge, Online (via Coursera)", "Apr 2026"),
      ...eduItem("Foundations of UX Design - Google, Online (via Coursera)", "2021", true),
    ]
  }]
});

function sectionHeading(text) {
  return new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: RULE, space: 1 } },
    spacing: { before: 52, after: 26 },
    children: [ new TextRun({ text, bold: true, size: 18, color: NAVY, allCaps: true }) ]
  });
}

function roleBlock(titleLine, dates, subLine, bullets, isLast) {
  const out = [];
  out.push(new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: RIGHT_TAB_POS }],
    spacing: { before: 52, after: 9 },
    children: [
      new TextRun({ text: titleLine, bold: true, size: 18 }),
      new TextRun({ text: `\t${dates}`, bold: true, size: 17 }),
    ]
  }));
  if (subLine) {
    out.push(new Paragraph({
      spacing: { after: 26 },
      children: [ new TextRun({ text: subLine, italics: true, size: 16, color: "404040" }) ]
    }));
  }
  bullets.forEach((b, i) => {
    out.push(new Paragraph({
      numbering: { reference: "bullets", level: 0 },
      spacing: { after: 13, ...LINE },
      children: [ new TextRun({ text: b, size: 18 }) ]
    }));
  });
  return out;
}

function skillLine(label, text, isLast) {
  return new Paragraph({
    spacing: { after: isLast ? 0 : 36, ...LINE },
    children: [
      new TextRun({ text: label + " ", bold: true, size: 18 }),
      new TextRun({ text, size: 18 }),
    ]
  });
}

function eduItem(titleText, dateText, isLast) {
  return [
    new Paragraph({
      tabStops: [{ type: TabStopType.RIGHT, position: RIGHT_TAB_POS }],
      spacing: { after: isLast ? 0 : 30 },
      children: [
        new TextRun({ text: titleText, size: 18, bold: true }),
        new TextRun({ text: `\t${dateText}`, size: 17, italics: true, color: "404040" }),
      ]
    })
  ];
}

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("Parth_Anand_Product_Designer_Resume.docx", buffer);
  console.log("done");
});
