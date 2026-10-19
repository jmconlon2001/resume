#!/usr/bin/env python3
"""
Build ATS-friendly versions of Jacob Conlon's resume from one source of truth.

Outputs (written next to this script):
  JacobConlon_Resume_ATS.docx  - upload this to Workday, iCIMS, Taleo, Greenhouse, Lever, etc.
  JacobConlon_Resume_ATS.txt   - plain text for "paste your resume" boxes
  WORKDAY_FIELDS.md            - field-by-field copy/paste sheet for when a parser still misses

Run:  python3 ats/build_ats_resume.py   (standard library only, no installs)

Why the .docx is built by hand: Workday's parser does best with a single-column
document that has no tables, text boxes, headers/footers, columns, images,
comments or tracked changes. Building the OOXML directly guarantees none of
that sneaks in.
"""
import os
import zipfile
from xml.sax.saxutils import escape

HERE = os.path.dirname(os.path.abspath(__file__))

# --------------------------------------------------------------------------
# CONTENT: edit here, then re-run the script.
# Dates: use MM/YYYY. Use "Present" for a current role.
# --------------------------------------------------------------------------
CONTACT = {
    "name": "Jacob Conlon",
    "city": "Cherry Hill", "state": "NJ", "zip": "08003",
    "phone": "856-996-5627",
    "email": "Jacob@Conlon.us",
    "linkedin": "https://www.linkedin.com/in/jacobconlon",
    "website": "https://www.conlon.us",
    "github": "https://github.com/jmconlon2001",
}

HEADLINE = "IT Director | Software Architecture, Cloud Modernization and AI Enablement"

SUMMARY = (
    "Technology leader with 19+ years building and modernizing enterprise software in "
    "financial services. Directs a 15-person team of architects, engineers, and QA, reporting "
    "to executive leadership. Combines deep technical expertise in .NET Core, Angular, cloud, "
    "and AI with forensic accounting fluency. Trusted to bridge IT and Finance on legacy "
    "migrations, GAAP-to-IFRS conversions, and enterprise AI adoption. Migrated a $2.2B lease "
    "portfolio (90,000 accounts, 290,000 active agreements) off a 25-year-old legacy platform "
    "during the global pandemic. IT owner for SOX and J-SOX controls."
)

JOBS = [
    {
        "title": "IT Director, Corporate IT",
        "company": "Canon Financial Services, Inc.",
        "city": "Mount Laurel", "state": "NJ",
        "start": "12/2019", "end": "Present",
        "bullets": [
            "Own IT application strategy, architecture, delivery, and platform support; lead a 15-person team of architects, engineers, and QA and report to executive leadership.",
            "Modeled and migrated a 25-year-old legacy lease accounting system during the global pandemic, a modernization peer institutions attempted and failed to complete. The platform carries a $2.2B lease portfolio across 90,000 accounts and 290,000 active agreements.",
            "Engineered a novel lease-syndication solution during a GAAP-to-IFRS accounting transition, eliminating a projected $1M project cost.",
            "Modernized the legacy green-screen leasing platform to a web and API based system; multi-year program delivered on time and on budget.",
            "As AI-assisted development champion, built ProcTuner, an AI agent that refactors underperforming SQL stored procedures, then validates data parity and benchmarks results before promoting. Some procedures gained more than 70% in performance.",
            "Driving cloud modernization and right-sizing: containerizing .NET Core services, migrating from Windows to Linux hosts, and adopting serverless deployments.",
            "Stood up application monitoring and observability that surfaces production issues in real time.",
            "Founded an IT Architecture Committee and instituted code-review standards, reducing defects reaching QA and strengthening security and performance.",
            "IT owner for SOX and J-SOX controls, embedding compliance into the secure SDLC.",
        ],
    },
    {
        "title": "Senior IT Manager, Applications and Systems",
        "company": "Canon Financial Services, Inc.",
        "city": "Mount Laurel", "state": "NJ",
        "start": "12/2016", "end": "12/2019",
        "bullets": [
            "Provided technical leadership for a team of 15 developers, architects, and QA.",
            "Launched a new short-term inventory-financing line of business on an external SaaS platform, increasing revenue and displacing competitor volume.",
            "Delivered a lease setup and booking application that cut manual data entry by more than 60%, lowered cost-to-serve, and created a framework for dealer data-feed integration.",
            "Replaced paper interdepartmental monetary requests with an electronic approval workflow using a dynamic authority matrix and automated posting to backend leasing systems.",
            "Built a Unified Desktop web application consolidating multiple internal and third-party systems, improving usability and reducing new-hire onboarding time.",
            "Implemented Salesforce Service Console integrated with telephony and backend data for call-center visibility and control.",
        ],
    },
    {
        "title": "Application Development Manager",
        "company": "Canon Financial Services, Inc.",
        "city": "Mount Laurel", "state": "NJ",
        "start": "12/2012", "end": "12/2016",
        "bullets": [
            "Mentored and led an 8-member development team in close partnership with Project Management.",
            "Built a risk-based credit decision engine that raised auto-approval rates from 8% to more than 50%.",
            "Delivered a customer self-service portal (payments, account servicing) integrated with the core lease-settlement system, raising payment-application rates and reducing inbound call volume.",
            "Shipped a multi-device dealer portal that increased credit-application volume from the independent dealer channel.",
            "Led gap analysis and systems integration enabling a successful merger with an acquired financial-services company.",
            "Introduced TDD and BDD practices (NUnit, SpecFlow) and Selenium automated UI testing.",
        ],
    },
    {
        "title": "Team Lead / Senior Engineer / Engineer",
        "company": "Canon Financial Services, Inc.",
        "city": "Mount Laurel", "state": "NJ",
        "start": "12/2007", "end": "12/2012",
        "bullets": [
            "Built a reusable .NET framework (Oracle and SQL Server data access, authentication, encryption, messaging, EDI, logging) adopted as the standard for all new applications.",
            "Developed prototype solutions with customers to accelerate requirements definition and development.",
        ],
    },
    {
        "title": "Senior Developer / Information Systems",
        "company": "Vision Appraisal Technology",
        "city": "Northboro", "state": "MA",
        "start": "09/2001", "end": "12/2007",
        "bullets": [
            "Owned full-lifecycle design and development of enterprise, database-driven applications for the Computer-Assisted Mass Appraisal (CAMA) industry.",
            "Served as technical project manager and lead developer for the New York City, Washington D.C., and Boston deployments.",
        ],
    },
]

EDUCATION = [
    {"school": "University of Phoenix", "degree": "Master of Science", "field": "Information Systems", "year": "2012"},
    {"school": "College of the Holy Cross", "degree": "Bachelor of Arts", "field": "History", "year": "2001"},
]

CERTIFICATIONS = [
    ("Certified Lease and Finance Professional (CLFP)", "CLFP Foundation", "2018"),
    ("Certified ScrumMaster (CSM)", "Scrum Alliance", "2010"),
    ("Six Sigma Green Belt", "Villanova University", "2010"),
    ("Microsoft Certified Application Developer (MCAD), C#", "Microsoft", "2006"),
    ("Canon Leadership Institute", "Canon U.S.A.", "2017"),
]

# Short, standard names so Workday can match them to its skills taxonomy.
SKILLS = [
    "IT Leadership", "Software Architecture", "Enterprise Architecture", "Solution Architecture",
    "Cloud Migration", "Legacy Modernization", "Team Leadership", "Vendor Management", "Budgeting",
    "C#", ".NET Core", "ASP.NET Web API", "Entity Framework", "LINQ", "Angular", "JavaScript",
    "jQuery", "HTML5", "REST APIs", "SOAP", "WCF", "Microservices",
    "Microsoft SQL Server", "Oracle Database", "T-SQL", "PL/SQL", "ETL", "Performance Tuning", "Data Modeling",
    "Amazon Web Services (AWS)", "Google Cloud Platform (GCP)", "Vertex AI", "Microsoft Azure", "Azure OpenAI",
    "Docker", "Azure DevOps", "CI/CD", "Linux", "F5 Load Balancing",
    "Artificial Intelligence (AI)", "Large Language Models (LLM)", "AI Agents", "GitHub Copilot", "Claude Code",
    "Salesforce", "EDI", "SaaS Integration", "Agile", "Scrum", "Test-Driven Development (TDD)",
    "Behavior-Driven Development (BDD)", "Selenium", "Code Review",
    "SOX Compliance", "J-SOX Compliance", "Secure SDLC", "Encryption",
    "Forensic Accounting", "IFRS", "GAAP", "Lease Accounting", "Financial Services",
]

# --------------------------------------------------------------------------
# Rendering helpers
# --------------------------------------------------------------------------
def location(d):
    return f"{d['city']}, {d['state']}"

def dates(j):
    return f"{j['start']} - {j['end']}"

def contact_lines():
    c = CONTACT
    return [
        f"{c['city']}, {c['state']} {c['zip']}",
        f"Phone: {c['phone']}",
        f"Email: {c['email']}",
        f"LinkedIn: {c['linkedin']}",
        f"Website: {c['website']}",
        f"GitHub: {c['github']}",
    ]

def cert_line(name, org, year):
    return ", ".join(p for p in (name, org, year) if p)

# --------------------------------------------------------------------------
# Plain text
# --------------------------------------------------------------------------
def build_txt():
    out = [CONTACT["name"].upper(), *contact_lines(), "", HEADLINE, "", "SUMMARY", SUMMARY, "", "WORK EXPERIENCE"]
    for j in JOBS:
        out += ["", j["title"], f"{j['company']}, {location(j)}", dates(j)]
        out += [f"- {b}" for b in j["bullets"]]
    out += ["", "EDUCATION"]
    for e in EDUCATION:
        out += [f"{e['degree']}, {e['field']}", e["school"], e["year"], ""]
    out += ["CERTIFICATIONS"] + [cert_line(*c) for c in CERTIFICATIONS]
    out += ["", "SKILLS", ", ".join(SKILLS), ""]
    return "\n".join(out)

# --------------------------------------------------------------------------
# Workday copy/paste sheet
# --------------------------------------------------------------------------
def build_fields_md():
    c = CONTACT
    md = [
        "# Workday Field-by-Field Sheet",
        "",
        "Generated by `build_ats_resume.py`. Do not edit by hand; edit the script and re-run it.",
        "",
        "Upload `JacobConlon_Resume_ATS.docx` first. Then check each field Workday filled in against this sheet,",
        "and copy/paste anything it missed. Each value is in a code block so it copies cleanly.",
        "",
        "## My Information",
        "",
        "| Field | Value |", "|---|---|",
        "| First Name | `Jacob` |", "| Last Name | `Conlon` |",
        f"| City | `{c['city']}` |", f"| State | `New Jersey` |", f"| Postal Code | `{c['zip']}` |",
        f"| Phone (Mobile) | `{c['phone']}` |", f"| Email | `{c['email']}` |",
        f"| LinkedIn | `{c['linkedin']}` |", f"| Website | `{c['website']}` |", f"| GitHub | `{c['github']}` |",
        "",
        "## Work Experience",
    ]
    for i, j in enumerate(JOBS, 1):
        current = j["end"] == "Present"
        md += [
            "", f"### {i}. {j['title']}", "",
            "| Field | Value |", "|---|---|",
            f"| Job Title | `{j['title']}` |",
            f"| Company | `{j['company']}` |",
            f"| Location | `{location(j)}` |",
            f"| I currently work here | {'Yes' if current else 'No'} |",
            f"| From | `{j['start']}` |",
        ]
        if not current:
            md.append(f"| To | `{j['end']}` |")
        md += ["", "Role Description:", "", "```", *[f"- {b}" for b in j["bullets"]], "```"]
    md += ["", "## Education"]
    for e in EDUCATION:
        md += [
            "", "| Field | Value |", "|---|---|",
            f"| School or University | `{e['school']}` |",
            f"| Degree | `{e['degree']}` |",
            f"| Field of Study | `{e['field']}` |",
            f"| To (Actual or Expected) | `{e['year']}` |",
        ]
    md += ["", "## Certifications", "", "| Certification | Issued By | Year |", "|---|---|---|"]
    md += [f"| `{n}` | {o} | {y or 'n/a'} |" for n, o, y in CERTIFICATIONS]
    md += [
        "", "## Skills", "",
        "Workday's skills box takes one skill at a time: type it, press Enter, then do the next one.",
        "", "```", *SKILLS, "```",
        "", "## Websites", "", "```", c["linkedin"], c["website"], c["github"], "```", "",
    ]
    return "\n".join(md)

# --------------------------------------------------------------------------
# DOCX (raw OOXML, single column, no tables/headers/text boxes)
# --------------------------------------------------------------------------
FONT = "Calibri"

def run(text, bold=False, size=None):
    rpr = ""
    if bold or size:
        rpr = "<w:rPr>" + ("<w:b/>" if bold else "") + (f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>' if size else "") + "</w:rPr>"
    return f'<w:r>{rpr}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'

def para(text="", style=None, bold=False, size=None, after=None):
    ppr = ""
    if style or after is not None:
        ppr = "<w:pPr>" + (f'<w:pStyle w:val="{style}"/>' if style else "") + (f'<w:spacing w:after="{after}"/>' if after is not None else "") + "</w:pPr>"
    return f"<w:p>{ppr}{run(text, bold, size) if text else ''}</w:p>"

def bullet(text):
    return ('<w:p><w:pPr><w:pStyle w:val="ListBullet"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>'
            + run(text) + "</w:p>")

def build_document_xml():
    b = []
    b.append(para(CONTACT["name"], style="Title"))
    for line in contact_lines():
        b.append(para(line, after=0))
    b.append(para(""))
    b.append(para(HEADLINE, bold=True))

    b.append(para("SUMMARY", style="Heading1"))
    b.append(para(SUMMARY))

    b.append(para("WORK EXPERIENCE", style="Heading1"))
    for j in JOBS:
        b.append(para(j["title"], style="Heading2"))
        b.append(para(f"{j['company']}, {location(j)}", after=0))
        b.append(para(dates(j), after=60))
        b += [bullet(x) for x in j["bullets"]]

    b.append(para("EDUCATION", style="Heading1"))
    for e in EDUCATION:
        b.append(para(f"{e['degree']}, {e['field']}", bold=True, after=0))
        b.append(para(e["school"], after=0))
        b.append(para(e["year"], after=120))

    b.append(para("CERTIFICATIONS", style="Heading1"))
    for c in CERTIFICATIONS:
        b.append(para(cert_line(*c), after=40))

    b.append(para("SKILLS", style="Heading1"))
    b.append(para(", ".join(SKILLS)))

    sect = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
            '<w:pgMar w:top="1008" w:right="1008" w:bottom="1008" w:left="1008" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f'<w:body>{"".join(b)}{sect}</w:body></w:document>')

STYLES_XML = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:eastAsia="{FONT}" w:cs="{FONT}"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="80" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
  <w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
    <w:pPr><w:spacing w:after="40"/></w:pPr><w:rPr><w:b/><w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:spacing w:before="240" w:after="80"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:spacing w:before="160" w:after="0"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="23"/><w:szCs w:val="23"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/><w:basedOn w:val="Normal"/><w:qFormat/>
    <w:pPr><w:spacing w:after="40"/><w:ind w:left="360" w:hanging="360"/></w:pPr></w:style>
</w:styles>'''

NUMBERING_XML = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="singleLevel"/>
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#8226;"/><w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}"/></w:rPr></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>'''

CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''

ROOT_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''

DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>'''

# Document properties carry the name and title; some parsers read these.
CORE_XML = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>{CONTACT["name"]} Resume</dc:title>
  <dc:creator>{CONTACT["name"]}</dc:creator>
  <cp:lastModifiedBy>{CONTACT["name"]}</cp:lastModifiedBy>
</cp:coreProperties>'''

APP_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Office Word</Application></Properties>'''

def build_docx(path):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("docProps/core.xml", CORE_XML)
        z.writestr("docProps/app.xml", APP_XML)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/document.xml", build_document_xml())
        z.writestr("word/styles.xml", STYLES_XML)
        z.writestr("word/numbering.xml", NUMBERING_XML)

if __name__ == "__main__":
    build_docx(os.path.join(HERE, "JacobConlon_Resume_ATS.docx"))
    with open(os.path.join(HERE, "JacobConlon_Resume_ATS.txt"), "w") as f:
        f.write(build_txt())
    with open(os.path.join(HERE, "WORKDAY_FIELDS.md"), "w") as f:
        f.write(build_fields_md())
    print("Built JacobConlon_Resume_ATS.docx, JacobConlon_Resume_ATS.txt, WORKDAY_FIELDS.md in", HERE)
