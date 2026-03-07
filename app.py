"""
CV Tailor — Production-ready
Tailors a CV to a job description using Claude via multi-round Q&A.
"""

import os
import re
import json
import uuid
import platform
import subprocess
import pdfplumber
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv

# Load .env from the same directory as this script
env_path = Path(__file__).parent / ".env"
load_dotenv(dotenv_path=env_path)

from flask import Flask, request, jsonify, send_file, render_template, redirect
from anthropic import Anthropic
from cv_generator import generate_cv_docx
import stripe

api_key = os.getenv("ANTHROPIC_API_KEY")
if not api_key:
    raise RuntimeError(
        "ANTHROPIC_API_KEY not found. "
        "Make sure .env file is in the same folder as app.py "
        f"(looked at: {env_path})"
    )

client = Anthropic(api_key=api_key)

stripe.api_key = os.getenv("STRIPE_SECRET_KEY")
STRIPE_PUBLISHABLE_KEY = os.getenv("STRIPE_PUBLISHABLE_KEY", "")
STRIPE_PRODUCT_ID = os.getenv("STRIPE_PRODUCT_ID", "")

# Model config: Sonnet for Q&A rounds (cheaper/faster), Opus for final generation
MODEL_ROUNDS = "claude-sonnet-4-6"
MODEL_FINAL = "claude-opus-4-6"

app = Flask(__name__)
app.secret_key = os.urandom(24)

# Force HTTPS in production
@app.before_request
def force_https():
    if request.headers.get('X-Forwarded-Proto', 'http') == 'http' and not request.host.startswith('127.0.0.1') and not request.host.startswith('localhost'):
        return redirect(request.url.replace('http://', 'https://'), code=301)

# In-memory session store
sessions = {}


# ---------------------------------------------------------------------------
# PDF CONVERSION (platform-aware)
# ---------------------------------------------------------------------------

def convert_docx_to_pdf(docx_path, pdf_path):
    """Convert .docx to .pdf using Word (Windows) or LibreOffice (Linux/Mac)."""
    if platform.system() == "Windows":
        try:
            import pythoncom
            from docx2pdf import convert as docx_to_pdf
            pythoncom.CoInitialize()
            docx_to_pdf(docx_path, pdf_path)
            pythoncom.CoUninitialize()
            return True
        except ImportError:
            print("[CV Tailor] docx2pdf not available, trying LibreOffice...")
    
    # Linux / Mac / fallback: use LibreOffice headless
    try:
        output_dir = str(Path(pdf_path).parent)
        result = subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "pdf",
             "--outdir", output_dir, str(docx_path)],
            capture_output=True, text=True, timeout=30
        )
        # LibreOffice names the output based on input filename
        expected_pdf = Path(output_dir) / (Path(docx_path).stem + ".pdf")
        if expected_pdf.exists() and str(expected_pdf) != pdf_path:
            expected_pdf.rename(pdf_path)
        return Path(pdf_path).exists()
    except (FileNotFoundError, subprocess.TimeoutExpired) as e:
        print(f"[CV Tailor] LibreOffice conversion failed: {e}")
        return False


# ---------------------------------------------------------------------------
# SYSTEM PROMPT
# ---------------------------------------------------------------------------

SYSTEM_PROMPT = """You are a senior recruiter and CV strategist with deep expertise in ATS screening systems, keyword optimization, and interview-stage candidate assessment. You have reviewed thousands of CVs and understand exactly what causes auto-discards and what makes a CV rise to the top of the pile.

YOUR MISSION: Tailor a candidate's CV to a specific job description through a structured 3-round interview, then generate an optimized single-page CV.

═══════════════════════════════════════════
STEP 1: ROLE-TYPE CLASSIFICATION
═══════════════════════════════════════════

Before anything else, classify the JD into one of these archetypes. This determines your entire strategy:

• CORPORATE OPERATIONS / STRATEGY — Budget ownership, cross-functional coordination, business planning, organizational enablement. EMPHASIZE: operational scope, budget figures, process improvement, stakeholder management, reporting/scorecards.

• CONSULTING / ADVISORY — Client-facing project delivery, proposal writing, team coaching, thought leadership. EMPHASIZE: project ownership, client deliverables, team management, BD contributions, industry expertise.

• INVESTMENT / FINANCE — Due diligence, financial modeling, deal execution, portfolio management. EMPHASIZE: deal volume, returns, fund size, IC presentations, sourcing metrics, board seats.

• COMMERCIAL / BD — Sales, partnerships, go-to-market, revenue generation. EMPHASIZE: revenue impact, pipeline metrics, client relationships, market expansion.

• TECHNICAL / R&D — Product development, engineering, scientific research. EMPHASIZE: technical skills, methodologies, publications, patents, system design.

• GENERAL MANAGEMENT — P&L ownership, team leadership, strategy + execution. EMPHASIZE: team size, revenue/budget scope, strategic initiatives, cross-functional impact.

The role type determines which of the candidate's experiences get MORE bullets (most relevant) vs FEWER bullets (supporting context), and what language register to use.

═══════════════════════════════════════════
STEP 2: GAP ANALYSIS METHODOLOGY
═══════════════════════════════════════════

Analyze the JD systematically:

A) MUST-HAVE REQUIREMENTS — Hard requirements that could cause auto-discard if missing. These include: years of experience, specific industry/sector, educational requirements, location/visa, explicit "required" skills.

B) KEYWORD GAPS — Terms and phrases the JD uses that don't appear anywhere in the CV. ATS systems often do keyword matching. Map each JD requirement to whether the CV addresses it.

C) QUANTIFICATION GAPS — Where the CV makes claims without numbers. Recruiters trust numbers: budget size, team size, deal count, revenue impact, project count, geographic scope.

D) FRAMING GAPS — Where the CV has relevant experience but describes it using the wrong language for this role type. Example: "managed investments" when the JD wants "led due diligence processes."

E) DIFFERENTIATION OPPORTUNITIES — What could make this candidate stand out, not just match.

Present your analysis summary as 3-5 concise bullet points (max 15 words each). Each bullet should be one clear finding. Do NOT write paragraphs.

═══════════════════════════════════════════
STEP 3: QUESTION STRATEGY (3 ROUNDS x 3 QUESTIONS)
═══════════════════════════════════════════

ROUND 1 — TARGET THE HIGHEST-RISK GAPS:
Ask about must-have requirements and auto-discard risks first. These are the gaps that could get the CV rejected before a human even reads it.

ROUND 2 — KEYWORD MATCHING & QUANTIFICATION:
With Round 1 answers incorporated, ask about remaining keyword gaps and unquantified claims. Focus on extracting numbers, metrics, and specific examples.
BEFORE asking Round 2 questions: infer the candidate's career trajectory and the narrative they're building with this application (e.g., "pivoting from VC into corporate strategy" or "deepening consulting expertise"). Use this narrative to frame questions that help them tell a coherent story.

ROUND 3 — DIFFERENTIATION & FINAL POLISH:
With Rounds 1-2 incorporated, ask about anything that could make the candidate stand out: unique achievements, cross-functional breadth, thought leadership, multicultural experience, etc.
BEFORE asking Round 3 questions: mentally draft the CV based on everything you know so far. Identify which bullet slots are still weak or empty, and ask questions that will directly fill those specific gaps.

HANDLING "I DON'T HAVE THIS EXPERIENCE" ANSWERS:
If a candidate says they lack experience with something, do NOT simply accept it and move on. In your next round, probe for ADJACENT or TRANSFERABLE experience that could partially address the gap. Most candidates underestimate their relevant experience. For example:
- "No direct budget oversight?" → Ask about involvement in budget planning, forecasting, cost tracking, or contributing to someone else's budget process.
- "No team management?" → Ask about mentoring, onboarding, coordinating with juniors, or leading project teams informally.
- "No industry experience?" → Ask about adjacent sectors, consulting/advisory work in the sector, or transferable domain knowledge.

QUESTION FORMAT RULES:
- Each question has 4 separate fields — keep them SHORT and distinct:
  - "question": The direct question (1-2 sentences MAX, specific, actionable)
  - "why": One sentence explaining what JD requirement this addresses
  - "example": A concrete example of what a strong answer looks like (one short sentence)
  - "targets": 2-5 word label of the JD requirement
- Keep the "question" field SHORT and direct. Do NOT merge the why/example into it.
- BAD question: "The JD requires budget oversight. At your current company, did you own a departmental budget? If so, what was the total amount and what did it cover including headcount and non-salary costs? A strong answer would include the total amount and your specific role."
- GOOD question: "Did you own a departmental budget? If so, what was the total amount and what did it cover?"
  why: "The JD requires direct budget oversight as a core accountability."
  example: "e.g. Managed €30M annual T&P budget covering salary and non-salary costs."
- Ask in plain prose, never bullet points or numbered lists

═══════════════════════════════════════════
STEP 4: FINAL CV GENERATION
═══════════════════════════════════════════

When generating the final CV:

STRUCTURE STRATEGY:
- Lead each role's bullets with the MOST JD-RELEVANT content, not chronological order
- Give MORE bullets to roles most relevant to the target JD, FEWER to supporting roles
- Add company descriptors in parentheses where the company isn't well-known (e.g., "Nina Capital (Healthcare VC Fund)")
- The skills "domain" line should contain ATS keywords matching the JD's terminology

LANGUAGE STRATEGY:
- Mirror the JD's exact terminology where truthful (if JD says "stakeholder management" don't write "working with people")
- Use strong action verbs: Led, Owned, Built, Managed, Drove, Spearheaded, Coordinated
- Every bullet should follow the pattern: [Bold Label]: [Action verb] + [what you did] + [scope/scale] + [impact/result]
- Keep bullet text concise — aim for 20-30 words per bullet. Prefer shorter. Every word must earn its place.

TRUTHFULNESS:
- NEVER fabricate experience. Only reframe and emphasize what the candidate has actually told you.
- If the candidate lacks a required qualification, note it in red_flags — do not invent it.

INPUT vs OUTPUT FORMAT:
- The input CV is ONLY a source of INFORMATION (work history, dates, companies, skills, achievements). Its formatting, structure, and organization are IRRELEVANT.
- The output CV ALWAYS follows a FIXED professional template regardless of how the input was formatted:
  1. Name (centered)
  2. Contact line (phone | email | location)
  3. PROFESSIONAL EXPERIENCE section with: Company + Location, Role Title + Dates, Bullet points with [Bold Label]: [Description]
  4. EDUCATION section with: Institution + Location, Degree + Dates
  5. SKILLS & EXPERTISE section with: Languages, Technical, Domain lines
- Even if the input CV is messy, unstructured, uses a different format, or has no clear sections — you MUST extract the relevant information and output it in the fixed template above.

═══════════════════════════════════════════
RESPONSE FORMATS (STRICT JSON — NO MARKDOWN)
═══════════════════════════════════════════

FOR ROUNDS 1-3, respond with ONLY this JSON:
{
  "round": 1,
  "role_type": "The archetype classification",
  "analysis": {
    "summary": ["3-5 concise bullet points assessing the CV against the JD. Each bullet: one clear finding, max 15 words."],
    "must_have_gaps": ["List of must-have JD requirements missing or weak in the CV"],
    "keyword_gaps": ["Specific JD keywords/phrases not in the CV"],
    "quantification_gaps": ["CV claims that lack numbers or specifics"],
    "strengths": ["What the CV already does well for this JD"]
  },
  "questions": [
    {
      "question": "Short direct question (1-2 sentences max)",
      "why": "One sentence: what JD requirement this addresses",
      "example": "e.g. A concrete example of a strong answer",
      "targets": "2-5 word JD requirement label"
    },
    {
      "question": "...",
      "why": "...",
      "example": "...",
      "targets": "..."
    },
    {
      "question": "...",
      "why": "...",
      "example": "...",
      "targets": "..."
    }
  ]
}

FOR THE FINAL OUTPUT (after round 3), respond with ONLY this JSON:
{
  "round": "final",
  "red_flags": ["Potential red flags — be specific and honest"],
  "cv": {
    "name": "Full Name",
    "contact": "phone | email | location",
    "experience": [
      {
        "company": "Company Name",
        "company_descriptor": "Optional parenthetical or null",
        "location": "City, Country",
        "roles": [
          {"title": "Job Title", "dates": "Start - End"}
        ],
        "bullets": [
          {"label": "Bold Label", "text": "Concise description (20-30 words max)"}
        ]
      }
    ],
    "education": [
      {
        "institution": "University Name",
        "location": "City, Country",
        "degree": "Degree description",
        "dates": "Start - End"
      }
    ],
    "skills": {
      "languages": "Language list with levels",
      "technical": "Technical skills with proficiency",
      "domain": "ATS-optimized domain keywords matching this JD"
    }
  }
}

CRITICAL: Respond with ONLY valid JSON. No markdown, no backticks, no preamble, no explanation outside the JSON structure."""


CONDENSE_SYSTEM = """You are a CV editor. You will receive a CV in JSON format that is too long to fit on a single page.

Your job is to make it fit by:
1. Shortening bullet text — cut filler words, combine overlapping bullets, tighten phrasing
2. Reducing bullets on less relevant roles (older or less relevant roles get 1-2 bullets max)
3. Removing redundant information that appears across multiple roles
4. Keeping the MOST impactful and JD-relevant content

RULES:
- Do NOT remove any roles or education entries entirely
- Do NOT change company names, dates, titles, or factual claims
- DO shorten bullet text aggressively — every word must earn its place
- DO merge or remove bullets that overlap with other bullets
- Target approximately 14-16 total bullets across all experience entries
- Each bullet's text should be 15-25 words maximum

Respond with ONLY the condensed CV JSON in the exact same schema. No markdown, no explanation."""


REVIEW_SYSTEM = """You are a CV quality reviewer. You will receive a tailored CV (JSON) and the original job description.

Your job is to improve the CV by:
1. Check each bullet — does it mirror the JD's language? Is it specific enough? Replace vague words with JD terminology.
2. Check the "domain" skills line — it must contain at least 5 keywords directly from the JD. Fix it if not.
3. Identify the single weakest bullet (least relevant to the JD) and either strengthen it or merge it into another bullet.
4. Ensure bullet labels are distinct — no two bullets should have similar labels.
5. Verify that the most JD-relevant experience leads each role's bullets.

RULES:
- Do NOT change company names, dates, titles, education, or factual claims.
- Do NOT add experience the candidate hasn't mentioned.
- DO improve word choice, keyword density, and bullet ordering.
- Keep the exact same JSON schema.

Respond with ONLY the improved CV JSON. No markdown, no explanation."""


# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------

def extract_pdf_text(filepath):
    """Extract text from a PDF file."""
    text = ""
    with pdfplumber.open(filepath) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    return text.strip()


def call_claude(messages, system=None, model=None):
    """Call Claude API and return the parsed JSON response."""
    response = client.messages.create(
        model=model or MODEL_ROUNDS,
        max_tokens=8192,
        system=system or SYSTEM_PROMPT,
        messages=messages,
    )
    raw = response.content[0].text.strip()
    if raw.startswith("```"):
        raw = raw.split("\n", 1)[1] if "\n" in raw else raw[3:]
        if raw.endswith("```"):
            raw = raw[:-3]
        raw = raw.strip()
    return json.loads(raw)


def pair_answers_with_questions(questions, answers):
    """Build answer text that pairs each answer with its original question."""
    parts = []
    for i, (q, a) in enumerate(zip(questions, answers)):
        if isinstance(q, dict):
            q_text = q.get("question", "")
        else:
            q_text = q
        parts.append(f'Question {i+1}: "{q_text}"\nAnswer {i+1}: {a}')
    return "\n\n".join(parts)


def generate_download_name(cv_data, jd_text):
    """Generate a dynamic filename like CV_FerranMarti_AZ_Mar2026.docx"""
    candidate_name = cv_data.get("name", "CV").replace(" ", "")
    jd_lines = jd_text.strip().split("\n")
    company = ""
    for line in jd_lines[:10]:
        clean = line.strip()
        if clean and len(clean) < 60 and not clean.lower().startswith(
            ("about", "we ", "the ", "our ", "join", "location", "deadline", "apply")
        ):
            company = clean
            break
    company = re.sub(r'[^a-zA-Z0-9 ]', '', company).strip()
    words = company.split()
    if len(words) >= 2:
        company_abbr = ''.join(w[0].upper() for w in words[:3])
    else:
        company_abbr = company[:3].upper()
    date_str = datetime.now().strftime("%b%Y")
    if company_abbr:
        return f"CV_{candidate_name}_{company_abbr}_{date_str}.docx"
    return f"CV_{candidate_name}_{date_str}.docx"


STOPWORDS = {
    "the", "a", "an", "and", "or", "but", "in", "on", "at", "to", "for", "of",
    "with", "by", "from", "as", "is", "was", "are", "were", "be", "been", "being",
    "have", "has", "had", "do", "does", "did", "will", "would", "could", "should",
    "may", "might", "shall", "can", "need", "must", "that", "this", "these", "those",
    "it", "its", "they", "them", "their", "we", "our", "you", "your", "he", "she",
    "his", "her", "all", "each", "every", "both", "few", "more", "most", "other",
    "some", "such", "than", "too", "very", "just", "also", "not", "only", "own",
    "same", "so", "about", "up", "out", "into", "over", "after", "before", "between",
    "through", "during", "above", "below", "under", "again", "further", "then",
    "once", "here", "there", "when", "where", "why", "how", "what", "which", "who",
    "whom", "any", "new", "well", "work", "working", "role", "including", "across",
    "ensure", "ability", "strong", "excellent", "required", "experience", "skills",
    "within", "part", "will", "including", "related",
}


def calculate_keyword_match(jd_text, cv_text):
    """Calculate percentage of JD keywords found in CV text."""
    jd_words = set(
        w.lower() for w in re.findall(r'\b[a-zA-Z]+\b', jd_text)
        if len(w) > 3 and w.lower() not in STOPWORDS
    )
    cv_words = set(
        w.lower() for w in re.findall(r'\b[a-zA-Z]+\b', cv_text)
    )
    if not jd_words:
        return 0, 0, set()
    matched = jd_words & cv_words
    return round(len(matched) / len(jd_words) * 100), len(jd_words), matched


def cv_json_to_text(cv_data):
    """Convert CV JSON to plain text for keyword matching."""
    parts = [cv_data.get("name", ""), cv_data.get("contact", "")]
    for exp in cv_data.get("experience", []):
        parts.append(exp.get("company", ""))
        parts.append(exp.get("company_descriptor", "") or "")
        for role in exp.get("roles", []):
            parts.append(role.get("title", ""))
        for bullet in exp.get("bullets", []):
            parts.append(bullet.get("label", ""))
            parts.append(bullet.get("text", ""))
    for edu in cv_data.get("education", []):
        parts.append(edu.get("degree", ""))
        parts.append(edu.get("institution", ""))
    skills = cv_data.get("skills", {})
    for key in ["languages", "technical", "domain"]:
        parts.append(skills.get(key, ""))
    return " ".join(parts)


# ---------------------------------------------------------------------------
# ROUTES
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/start", methods=["POST"])
def start():
    """Upload CV + JD, kick off round 1."""
    jd_text = request.form.get("jd", "").strip()
    cv_file = request.files.get("cv")

    if not jd_text or not cv_file:
        return jsonify({"error": "Please provide both a CV file and job description."}), 400

    sid = str(uuid.uuid4())
    upload_dir = os.path.join("outputs", sid)
    os.makedirs(upload_dir, exist_ok=True)
    pdf_path = os.path.join(upload_dir, "original_cv.pdf")
    cv_file.save(pdf_path)

    cv_text = extract_pdf_text(pdf_path)
    if not cv_text:
        return jsonify({"error": "Could not extract text from the PDF."}), 400

    user_msg = f"""Here is the candidate's current CV:

<cv>
{cv_text}
</cv>

Here is the job description they want to tailor their CV for:

<jd>
{jd_text}
</jd>

INSTRUCTIONS:
1. Classify this role into one of the archetypes from your framework.
2. Perform the full structured gap analysis.
3. Ask your Round 1 questions — targeting the highest-risk gaps.

Each question must have separate "question", "why", "example", and "targets" fields.

Respond with the Round 1 JSON."""

    messages = [{"role": "user", "content": user_msg}]

    try:
        result = call_claude(messages, model=MODEL_ROUNDS)
    except Exception as e:
        return jsonify({"error": f"Claude API error: {str(e)}"}), 500

    messages.append({"role": "assistant", "content": json.dumps(result)})
    sessions[sid] = {
        "messages": messages,
        "round": 1,
        "output_dir": upload_dir,
        "jd_text": jd_text,
        "cv_text": cv_text,
        "last_questions": result.get("questions", []),
    }

    return jsonify({"session_id": sid, **result})


@app.route("/answer", methods=["POST"])
def answer():
    """Submit answers for current round, get next round or final CV."""
    data = request.get_json()
    sid = data.get("session_id")
    answers = data.get("answers", [])

    if sid not in sessions:
        return jsonify({"error": "Session not found."}), 404

    sess = sessions[sid]
    current_round = sess["round"]

    paired_text = pair_answers_with_questions(
        sess.get("last_questions", []),
        answers
    )

    if current_round < 3:
        round_specific = ""
        if current_round == 1:
            round_specific = """- Round 2 should focus on KEYWORD MATCHING and QUANTIFICATION gaps — extract specific numbers, metrics, examples that are missing.
- Before asking questions, infer the candidate's career trajectory — what story are they telling with this application? Use this narrative to frame questions that help them build a coherent arc."""
        else:
            round_specific = """- Round 3 should focus on DIFFERENTIATION and FINAL POLISH — what makes this candidate stand out, cross-functional breadth, thought leadership, multicultural experience.
- Before asking questions, mentally draft the CV based on everything you know. Identify which bullet slots are still weak or empty. Ask questions that will directly fill those gaps."""

        user_msg = f"""Here are the candidate's answers to Round {current_round}, paired with the original questions for clarity:

{paired_text}

INSTRUCTIONS:
1. Reassess your gap analysis with this new information. What gaps are now covered? What remains?
2. If the candidate said they lack experience with something, probe for ADJACENT or TRANSFERABLE experience in this round.
3. Update the structured analysis (must_have_gaps, keyword_gaps, quantification_gaps, strengths).
4. Ask 3 NEW questions for Round {current_round + 1}:
   {round_specific}
5. Each question must have separate "question", "why", "example", and "targets" fields.

Respond with the Round {current_round + 1} JSON."""
    else:
        user_msg = f"""Here are the candidate's answers to Round 3 (final round), paired with the original questions:

{paired_text}

NOW GENERATE THE FINAL TAILORED CV.

Review everything the candidate has shared across all 3 rounds. Then build the CV following these rules:

STRUCTURE:
- Lead each role's bullets with the MOST JD-RELEVANT content first
- Give MORE bullets to roles most relevant to this specific JD, FEWER to supporting roles
- Add company descriptors in parentheses where the company name alone isn't self-explanatory
- Order experience entries reverse-chronologically (most recent first)

LANGUAGE:
- Mirror the JD's exact terminology where truthful
- Use strong action verbs: Led, Owned, Built, Managed, Drove, Spearheaded, Coordinated
- Pattern: [Bold Label]: [Action verb] + [what] + [scope/scale] + [impact]
- Keep each bullet's text to 20-30 words. Be concise. Every word earns its place.

SKILLS SECTION:
- "domain" line must contain ATS-optimized keywords pulled directly from this JD's language

TRUTHFULNESS:
- NEVER fabricate. Only reframe what the candidate has actually shared.
- Genuine gaps go in red_flags.

Respond with the final JSON."""

    sess["messages"].append({"role": "user", "content": user_msg})

    # Use Sonnet for rounds, Opus for final generation
    use_model = MODEL_FINAL if current_round >= 3 else MODEL_ROUNDS

    try:
        result = call_claude(sess["messages"], model=use_model)
    except Exception as e:
        return jsonify({"error": f"Claude API error: {str(e)}"}), 500

    sess["messages"].append({"role": "assistant", "content": json.dumps(result)})

    if current_round < 3:
        sess["round"] = current_round + 1
        sess["last_questions"] = result.get("questions", [])
        return jsonify({"session_id": sid, **result})
    else:
        # Final round — generate the docx, check page count via PDF
        try:
            cv_data = result["cv"]
            output_dir = sess["output_dir"]
            docx_path = os.path.join(output_dir, "tailored_cv.docx")
            pdf_path = os.path.join(output_dir, "tailored_cv.pdf")

            attempts = 0
            while attempts < 3:
                generate_cv_docx(cv_data, docx_path)

                converted = convert_docx_to_pdf(docx_path, pdf_path)
                if not converted:
                    print("[CV Tailor] PDF conversion unavailable — skipping page check.")
                    break

                with pdfplumber.open(pdf_path) as pdf:
                    num_pages = len(pdf.pages)
                    print(f"[CV Tailor] PDF pages: {num_pages}")

                    if num_pages <= 1:
                        break

                    overflow_text = ""
                    for page in pdf.pages[1:]:
                        page_text = page.extract_text()
                        if page_text:
                            overflow_text += page_text

                attempts += 1
                print(f"[CV Tailor] Overflows to {num_pages} pages — condensing (attempt {attempts})...")

                trim_msg = [{
                    "role": "user",
                    "content": f"""This CV overflows to {num_pages} pages. The following content spilled onto page 2:

---OVERFLOW TEXT---
{overflow_text}
---END---

Make the CV more concise so EVERYTHING fits on a single page. Shorten bullet text, merge overlapping bullets, trim the least important details. Do NOT remove any roles or education entries.

<cv_json>
{json.dumps(cv_data, indent=2)}
</cv_json>

Return ONLY the trimmed CV JSON in the exact same schema."""
                }]
                cv_data = call_claude(trim_msg, system=CONDENSE_SYSTEM, model=MODEL_FINAL)

            sess["docx_path"] = docx_path

            # Quality review pass (point 4) — one cheap Sonnet call
            print("[CV Tailor] Running quality review pass...")
            review_msg = [{
                "role": "user",
                "content": f"""Review this tailored CV against the original job description. Improve keyword density, strengthen weak bullets, and ensure the domain skills line contains JD keywords.

<jd>
{sess["jd_text"][:3000]}
</jd>

<cv_json>
{json.dumps(cv_data, indent=2)}
</cv_json>

Return ONLY the improved CV JSON."""
            }]
            try:
                cv_data = call_claude(review_msg, system=REVIEW_SYSTEM, model=MODEL_ROUNDS)
                # Regenerate docx with improved version
                generate_cv_docx(cv_data, docx_path)
                convert_docx_to_pdf(docx_path, pdf_path)
                print("[CV Tailor] Quality review complete.")
            except Exception as review_err:
                print(f"[CV Tailor] Quality review failed (using pre-review version): {review_err}")

            # Calculate keyword match scores (point 9)
            original_score, total_keywords, _ = calculate_keyword_match(
                sess["jd_text"], sess["cv_text"]
            )
            tailored_text = cv_json_to_text(cv_data)
            tailored_score, _, _ = calculate_keyword_match(
                sess["jd_text"], tailored_text
            )
            print(f"[CV Tailor] Keyword match: {original_score}% → {tailored_score}% ({total_keywords} JD keywords)")

            # Store PDF path for preview
            sess["pdf_path"] = pdf_path

            download_name = generate_download_name(cv_data, sess["jd_text"])
            sess["download_name"] = download_name

            return jsonify({
                "session_id": sid,
                "round": "final",
                "red_flags": result.get("red_flags", []),
                "download_ready": True,
                "download_name": download_name,
                "keyword_match_original": original_score,
                "keyword_match_tailored": tailored_score,
                "total_keywords": total_keywords,
            })
        except Exception as e:
            return jsonify({"error": f"CV generation error: {str(e)}"}), 500


@app.route("/download/<sid>")
def download(sid):
    """Download the generated CV — only if paid."""
    if sid not in sessions or "docx_path" not in sessions[sid]:
        return jsonify({"error": "No CV available for download."}), 404
    if not sessions[sid].get("paid"):
        return jsonify({"error": "Payment required."}), 402
    return send_file(
        sessions[sid]["docx_path"],
        as_attachment=True,
        download_name=sessions[sid].get("download_name", "tailored_cv.docx"),
    )


@app.route("/preview-original/<sid>")
def preview_original(sid):
    """Serve the original CV as a PNG image."""
    if sid not in sessions:
        return jsonify({"error": "Session not found."}), 404
    original_path = os.path.join(sessions[sid]["output_dir"], "original_cv.pdf")
    img_path = os.path.join(sessions[sid]["output_dir"], "original_preview.png")
    if not os.path.exists(img_path):
        try:
            with pdfplumber.open(original_path) as pdf:
                page = pdf.pages[0]
                img = page.to_image(resolution=120)
                img.save(img_path)
        except Exception as e:
            return jsonify({"error": f"Preview failed: {str(e)}"}), 500
    return send_file(img_path, mimetype="image/png")


@app.route("/preview-tailored/<sid>")
def preview_tailored(sid):
    """Serve the tailored CV as a PNG image."""
    if sid not in sessions or "pdf_path" not in sessions[sid]:
        return jsonify({"error": "Tailored CV not found."}), 404
    pdf_path = sessions[sid]["pdf_path"]
    img_path = os.path.join(sessions[sid]["output_dir"], "tailored_preview.png")
    if not os.path.exists(img_path):
        try:
            with pdfplumber.open(pdf_path) as pdf:
                page = pdf.pages[0]
                img = page.to_image(resolution=120)
                img.save(img_path)
        except Exception as e:
            return jsonify({"error": f"Preview failed: {str(e)}"}), 500
    return send_file(img_path, mimetype="image/png")


@app.route("/create-checkout", methods=["POST"])
def create_checkout():
    """Create a Stripe Checkout session for €0.99."""
    data = request.get_json()
    sid = data.get("session_id")

    if not sid or sid not in sessions:
        return jsonify({"error": "Session not found."}), 404

    try:
        checkout_session = stripe.checkout.Session.create(
            payment_method_types=["card"],
            line_items=[{
                "price_data": {
                    "currency": "eur",
                    "product": STRIPE_PRODUCT_ID,
                    "unit_amount": 399,
                },
                "quantity": 1,
            }],
            mode="payment",
            success_url=request.host_url + f"payment-success?session_id={sid}&stripe_session={{CHECKOUT_SESSION_ID}}",
            cancel_url=request.host_url + f"payment-cancel?session_id={sid}",
            metadata={"cv_session_id": sid},
        )
        return jsonify({"checkout_url": checkout_session.url})
    except Exception as e:
        return jsonify({"error": f"Stripe error: {str(e)}"}), 500


@app.route("/payment-success")
def payment_success():
    """Stripe redirects here after successful payment."""
    sid = request.args.get("session_id")
    stripe_session_id = request.args.get("stripe_session")

    if not sid or sid not in sessions:
        return render_template("payment_result.html", success=False, error="Session not found.")

    try:
        checkout = stripe.checkout.Session.retrieve(stripe_session_id)
        if checkout.payment_status == "paid":
            sessions[sid]["paid"] = True
            return render_template("payment_result.html",
                success=True,
                session_id=sid,
                download_name=sessions[sid].get("download_name", "tailored_cv.docx"))
        else:
            return render_template("payment_result.html", success=False, error="Payment not completed.")
    except Exception as e:
        return render_template("payment_result.html", success=False, error=str(e))


@app.route("/payment-cancel")
def payment_cancel():
    """User cancelled payment."""
    sid = request.args.get("session_id")
    return render_template("payment_result.html", success=False, error="Payment was cancelled. You can try again from the app.")


@app.route("/retry", methods=["POST"])
def retry():
    """Start a new tailoring session with the same CV but a new JD."""
    data = request.get_json()
    old_sid = data.get("session_id")
    jd_text = data.get("jd", "").strip()

    if not old_sid or old_sid not in sessions:
        return jsonify({"error": "Previous session not found."}), 404
    if not jd_text:
        return jsonify({"error": "Please provide a job description."}), 400

    cv_text = sessions[old_sid].get("cv_text")
    if not cv_text:
        return jsonify({"error": "CV text not found in previous session."}), 404

    sid = str(uuid.uuid4())
    upload_dir = os.path.join("outputs", sid)
    os.makedirs(upload_dir, exist_ok=True)

    user_msg = f"""Here is the candidate's current CV:

<cv>
{cv_text}
</cv>

Here is the job description they want to tailor their CV for:

<jd>
{jd_text}
</jd>

INSTRUCTIONS:
1. Classify this role into one of the archetypes from your framework.
2. Perform the full structured gap analysis.
3. Ask your Round 1 questions — targeting the highest-risk gaps.

Each question must have separate "question", "why", "example", and "targets" fields.

Respond with the Round 1 JSON."""

    messages = [{"role": "user", "content": user_msg}]

    try:
        result = call_claude(messages, model=MODEL_ROUNDS)
    except Exception as e:
        return jsonify({"error": f"Claude API error: {str(e)}"}), 500

    messages.append({"role": "assistant", "content": json.dumps(result)})
    sessions[sid] = {
        "messages": messages,
        "round": 1,
        "output_dir": upload_dir,
        "jd_text": jd_text,
        "cv_text": cv_text,
        "last_questions": result.get("questions", []),
    }

    return jsonify({"session_id": sid, **result})


if __name__ == "__main__":
    app.run(debug=True, port=5050)
