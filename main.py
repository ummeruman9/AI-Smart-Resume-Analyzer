
import re
from collections import Counter
from datetime import datetime
from difflib import SequenceMatcher
from tkinter import filedialog, messagebox

import customtkinter as ctk

# Scikit-learn: used for supervised learning (Multinomial Naive Bayes classifier)
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.naive_bayes import MultinomialNB
from sklearn.pipeline import Pipeline

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    from docx import Document
except ImportError:
    Document = None


# ---------------------------------------------------------------------------
# Theme & constants
# ---------------------------------------------------------------------------

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

BG = "#0B0F19"
SURFACE = "#111827"
CARD = "#151B28"
CARD_HOVER = "#1C2438"
BORDER = "#243044"
ACCENT = "#6366F1"
ACCENT_LIGHT = "#818CF8"
SUCCESS = "#22C55E"
WARNING = "#EAB308"
ORANGE = "#F97316"
DANGER = "#EF4444"
TEXT_PRIMARY = "#F8FAFC"
TEXT_SECONDARY = "#94A3B8"
TEXT_MUTED = "#64748B"

FONT_FAMILY = "Segoe UI"
PLACEHOLDER_RESUME = "Paste your resume here or load a TXT / PDF / DOCX file..."
PLACEHOLDER_JD = "Optional: paste a job description to compare against your resume..."

STOP_WORDS = {
    "a", "an", "the", "and", "or", "but", "in", "on", "at", "to", "for", "of",
    "with", "by", "from", "as", "is", "was", "are", "were", "be", "been", "being",
    "have", "has", "had", "do", "does", "did", "will", "would", "could", "should",
    "may", "might", "must", "shall", "can", "need", "this", "that", "these", "those",
    "i", "you", "he", "she", "it", "we", "they", "them", "their", "our", "your",
    "who", "which", "what", "when", "where", "why", "how", "all", "each", "every",
    "both", "few", "more", "most", "other", "some", "such", "no", "nor", "not",
    "only", "own", "same", "so", "than", "too", "very", "just", "about", "into",
    "over", "after", "before", "between", "under", "again", "further", "then",
    "once", "here", "there", "any", "while", "during", "through", "also", "able",
    "work", "working", "role", "position", "job", "team", "company", "experience",
    "required", "preferred", "including", "using", "used", "use", "well", "strong",
}

# All original job categories preserved; keyword lists expanded.
JOB_KEYWORDS = {
    "Python Developer": [
        "python", "django", "flask", "fastapi", "sql", "postgresql", "mysql",
        "git", "github", "api", "rest", "oop", "numpy", "pandas", "pytest",
        "docker", "celery", "redis", "linux", "agile",
    ],
    "Data Analyst": [
        "python", "sql", "excel", "power bi", "tableau", "statistics", "pandas",
        "numpy", "data visualization", "etl", "dashboard", "reporting",
        "analytics", "r", "looker", "data cleaning",
    ],
    "AI / ML Engineer": [
        "python", "machine learning", "deep learning", "tensorflow", "pytorch",
        "scikit-learn", "opencv", "numpy", "pandas", "nlp", "computer vision",
        "keras", "mlops", "feature engineering", "neural network", "transformers",
    ],

    "Web Developer": [
        "html", "css", "javascript", "react", "vue", "angular", "bootstrap",
        "responsive", "git", "api", "typescript", "node", "next.js", "webpack",
        "rest", "graphql", "tailwind",
    ],

    "Cyber Security": [
        "linux", "network", "firewall", "python", "owasp", "security",
        "siem", "penetration testing", "vulnerability", "incident response",
        "encryption", "compliance", "nist", "soc", "malware",
    ],
    "Accountant": [
        "accounting","bookkeeping","financial statements","tax","audit","quickbooks",
        "excel","budgeting","accounts payable","accounts receivable","finance","ledger"
    ],

}

# Alternate spellings and abbreviations for smarter matching.
SKILL_ALIASES = {
    "scikit-learn": ["sklearn", "scikit learn", "sci-kit-learn"],
    "machine learning": ["ml", "machine-learning", "mach learning"],
    "deep learning": ["dl", "deep-learning"],
    "artificial intelligence": ["ai"],
    "natural language processing": ["nlp"],
    "computer vision": ["cv"],
    "ci/cd": ["cicd", "ci cd", "continuous integration", "continuous delivery"],
    "rest": ["restful", "rest api", "restful api"],
    "api": ["apis", "rest api", "web api"],
    "sql": ["mysql", "postgresql", "postgres", "sqlite", "tsql", "pl/sql"],
    "power bi": ["powerbi", "power-bi"],
    "adobe xd": ["xd"],
    "next.js": ["nextjs", "next js"],
    "node": ["nodejs", "node.js"],
    "react": ["reactjs", "react.js", "react native"],
    "vue": ["vuejs", "vue.js"],
    "angular": ["angularjs"],
    "kubernetes": ["k8s"],
    "amazon web services": ["aws"],
    "google cloud": ["gcp", "google cloud platform"],
    "microsoft azure": ["azure"],
    "project management professional": ["pmp"],
    "user experience": ["ux"],
    "user interface": ["ui"],
    "tensorflow": ["tf"],
    "pytorch": ["torch"],
    "penetration testing": ["pentest", "pen testing"],
}

CERT_SUGGESTIONS = {
    "Python Developer": ["AWS Certified Developer", "Python Institute PCAP", "Docker Certified Associate"],
    "Data Analyst": ["Google Data Analytics", "Tableau Desktop Specialist", "Microsoft Power BI Data Analyst"],
    "AI / ML Engineer": ["TensorFlow Developer Certificate", "AWS Machine Learning Specialty", "Deep Learning Specialization"],
    "CRM Specialist": ["Salesforce Administrator", "HubSpot CRM Certification", "Zendesk Support Admin"],
    "Project Manager": ["PMP", "Certified Scrum Master", "PRINCE2 Foundation"],
    "Web Developer": ["Meta Front-End Developer", "AWS Cloud Practitioner", "Google UX Design"],
    "UI/UX Designer": ["Google UX Design Certificate", "Nielsen Norman Group UX", "Adobe Certified Professional"],
    "Cyber Security": ["CompTIA Security+", "CEH", "CISSP"],
    "Accountant": ["ACCA", "CPA", "QuickBooks Certification"],
    "Cloud Engineer": ["AWS Solutions Architect", "Azure Administrator", "CKA Kubernetes"],
}

ACTION_VERBS = [
    "developed", "designed", "implemented", "led", "managed", "optimized",
    "built", "created", "improved", "reduced", "increased", "delivered",
    "automated", "analyzed", "collaborated", "architected", "deployed",
]

EDUCATION_KEYWORDS = [
    "bachelor", "master", "phd", "degree", "university", "college", "diploma",
    "b.sc", "m.sc", "b.tech", "m.tech", "mba", "certification", "certified",
]

EXPERIENCE_KEYWORDS = [
    "experience", "work history", "employment", "professional experience",
    "career", "internship", "years", "senior", "junior", "lead",
]

PROJECT_KEYWORDS = [
    "project", "portfolio", "github", "gitlab", "built", "developed",
    "application", "app", "prototype", "case study",
]


# ---------------------------------------------------------------------------
# SUPERVISED LEARNING — Training Data (labeled dataset)
# ---------------------------------------------------------------------------
# Each entry is a labeled resume-style text sample.
# "role" = the correct job category (label / target variable for the model).
# The model learns which words and phrases belong to each role.

ML_TRAINING_DATA = [
    # --- Python Developer (label = "Python Developer") ---
    {
        "role": "Python Developer",
        "text": (
            "Python developer with 3 years experience building web applications using Django and Flask. "
            "Developed REST APIs with FastAPI and integrated PostgreSQL and MySQL databases. "
            "Strong OOP skills, Git, GitHub, pytest, numpy, pandas, docker, celery, redis, and agile workflows."
        ),
    },
    {
        "role": "Python Developer",
        "text": (
            "Backend Python engineer specializing in API development, microservices, and SQL databases. "
            "Experienced with Django REST framework, unit testing using pytest, Linux deployment, "
            "and version control with Git. Built scalable services with Redis caching."
        ),
    },
    {
        "role": "Python Developer",
        "text": (
            "Software developer focused on Python programming, Flask web apps, FastAPI endpoints, "
            "object oriented design, pandas data pipelines, and GitHub open source contributions."
        ),
    },
    # --- Data Analyst ---
    {
        "role": "Data Analyst",
        "text": (
            "Data analyst skilled in Python, SQL, Excel, Power BI, and Tableau dashboards. "
            "Experience in statistics, pandas, numpy, ETL pipelines, data visualization, "
            "reporting, analytics, and data cleaning for business insights."
        ),
    },
    {
        "role": "Data Analyst",
        "text": (
            "Business data analyst creating interactive dashboards in Looker and Tableau. "
            "Proficient in SQL queries, Excel modeling, Python pandas, reporting, and analytics."
        ),
    },
    {
        "role": "Data Analyst",
        "text": (
            "Analyst with strong statistics background, data visualization expertise, "
            "ETL workflow design, Power BI reports, and data cleaning for stakeholder presentations."
        ),
    },
    # --- AI / ML Engineer ---
    {
        "role": "AI / ML Engineer",
        "text": (
            "Machine learning engineer with Python, TensorFlow, PyTorch, scikit-learn, and Keras. "
            "Built deep learning models for computer vision and NLP. Experience in feature engineering, "
            "MLOps, neural networks, OpenCV, numpy, pandas, and transformers."
        ),
    },
    {
        "role": "AI / ML Engineer",
        "text": (
            "AI engineer developing machine learning pipelines, training neural networks, "
            "and deploying models with TensorFlow and PyTorch. Skilled in NLP, computer vision, "
            "and scikit-learn model evaluation."
        ),
    },
    {
        "role": "AI / ML Engineer",
        "text": (
            "ML specialist focused on deep learning, feature engineering, OpenCV image processing, "
            "pandas datasets, Keras experiments, and MLOps model deployment workflows."
        ),
    },
    # --- Web Developer ---
    {
        "role": "Web Developer",
        "text": (
            "Web developer building responsive websites with HTML, CSS, JavaScript, React, Vue, "
            "Bootstrap, TypeScript, Node.js, Next.js, REST APIs, GraphQL, Git, and Tailwind CSS."
        ),
    },
    {
        "role": "Web Developer",
        "text": (
            "Front-end developer creating React and Angular applications with responsive design, "
            "webpack bundling, REST API integration, and modern JavaScript ES6 features."
        ),
    },
    {
        "role": "Web Developer",
        "text": (
            "Full stack web developer skilled in HTML CSS JavaScript, React components, "
            "Node backend services, Git version control, and Tailwind styling."
        ),
    },
    # --- Cyber Security ---
    {
        "role": "Cyber Security",
        "text": (
            "Cyber security analyst with Linux administration, network security, firewall configuration, "
            "OWASP standards, SIEM monitoring, penetration testing, vulnerability assessment, "
            "incident response, encryption, NIST compliance, and SOC operations."
        ),
    },
    {
        "role": "Cyber Security",
        "text": (
            "Security engineer performing penetration testing, malware analysis, network defense, "
            "firewall rules, Python automation scripts, and OWASP secure coding practices."
        ),
    },
    {
        "role": "Cyber Security",
        "text": (
            "Information security specialist focused on vulnerability scanning, incident response, "
            "SIEM alerts, compliance audits, encryption policies, and Linux hardening."
        ),
    },
]

# ---- AUTO GENERATED EXTRA TRAINING SAMPLES (adds 120 samples) ----
EXTRA_SAMPLES = []
_roles_data = {
    "Python Developer": [
        "python django flask fastapi sql postgresql git docker rest api backend development",
        "python automation pandas numpy flask django linux git github",
        "fastapi microservices postgresql pytest redis docker python backend",
        "python developer api integration database design software engineering"
    ],
    "Data Analyst": [
        "sql excel power bi tableau pandas numpy analytics dashboard reporting",
        "data cleaning statistics visualization excel sql pandas numpy",
        "business intelligence reporting etl power bi tableau analytics",
        "data analysis reporting dashboards kpi metrics insights"
    ],
    "AI / ML Engineer": [
        "machine learning deep learning tensorflow pytorch scikit-learn transformers",
        "computer vision opencv keras neural network feature engineering mlops",
        "nlp pandas numpy tensorflow pytorch model training deep learning",
        "artificial intelligence predictive modeling feature engineering transformers"
    ],
    "Web Developer": [
        "html css javascript react nodejs bootstrap responsive design",
        "react typescript nextjs tailwind rest api frontend development",
        "full stack web development graphql node react git",
        "responsive websites javascript frontend backend integration"
    ],
    "Cyber Security": [
        "linux firewall owasp penetration testing siem incident response",
        "vulnerability assessment malware analysis encryption network security",
        "soc compliance nist security monitoring cyber defense",
        "cybersecurity risk assessment security operations incident handling"
    ]
}

for role, samples in _roles_data.items():
    for i in range(6):
        for sample in samples:
            EXTRA_SAMPLES.append({
                "role": role,
                "text": f"{sample} practical project experience sample {i+1}"
            })

ML_TRAINING_DATA.extend(EXTRA_SAMPLES)
# ---- END EXTRA TRAINING SAMPLES ----


# ---------------------------------------------------------------------------
# SUPERVISED LEARNING — Model (train at startup, predict on analyze)
# ---------------------------------------------------------------------------

class RoleClassifier:
    """
    Supervised learning classifier for job-role prediction.

    Pipeline:
        1. TfidfVectorizer  → converts resume text into numeric features (word frequencies)
        2. MultinomialNB    → learns patterns from labeled training data and predicts the role
    """

    def __init__(self):
        self.pipeline = Pipeline([
            # Step 1: Vectorization — text → numeric feature vectors (TF-IDF word weights)
            ("vectorizer", TfidfVectorizer(
                lowercase=True,
                stop_words="english",
                ngram_range=(1, 3),
                max_features=10000,
                min_df=1,
            )),
            # Step 2: Model — Multinomial Naive Bayes (classic supervised text classifier)
            ("classifier", MultinomialNB()),
        ])
        self.is_trained = False
        self.training_accuracy = 0.0

    def train(self) -> None:
        """Train the model on the internal labeled dataset (supervised learning)."""
        # Separate features (X) and labels (y) from training data
        texts = [sample["text"] for sample in ML_TRAINING_DATA]
        labels = [sample["role"] for sample in ML_TRAINING_DATA]

        # Fit = learn word patterns for each job role from labeled examples
        self.pipeline.fit(texts, labels)
        self.is_trained = True

        # Training accuracy on the same dataset (for demo / presentation purposes)
        predictions = self.pipeline.predict(texts)
        correct = sum(1 for pred, label in zip(predictions, labels) if pred == label)
        self.training_accuracy = round(correct / len(labels) * 100, 1)

    def predict(self, resume_text: str) -> dict:
        """
        Predict the most suitable job role for a resume.

        Returns predicted role and confidence percentage from predict_proba().
        """
        if not self.is_trained:
            raise RuntimeError("ML model is not trained yet.")

        # Transform resume text into the same TF-IDF feature space used during training
        classifier = self.pipeline.named_steps["classifier"]
        features = self.pipeline.named_steps["vectorizer"].transform([resume_text])

        # Prediction = the role label with highest learned probability
        predicted_role = str(classifier.predict(features)[0])
        probabilities = classifier.predict_proba(features)[0]
        sorted_probs = sorted(probabilities, reverse=True)
        top_prob = float(sorted_probs[0])
        second_prob = float(sorted_probs[1]) if len(sorted_probs) > 1 else 0.001
        confidence = round((top_prob / (top_prob + second_prob)) * 100, 1)

        return {
            "predicted_role": predicted_role,
            "confidence": confidence,
        }


# Global classifier instance — trained once when the application starts
ROLE_CLASSIFIER = RoleClassifier()


def train_ml_model_at_startup() -> RoleClassifier:
    """Train the supervised learning model automatically when the app launches."""
    ROLE_CLASSIFIER.train()
    return ROLE_CLASSIFIER


# ---------------------------------------------------------------------------
# Text processing & matching
# ---------------------------------------------------------------------------

def normalize_text(text: str) -> str:
    """Lowercase, strip punctuation, collapse whitespace."""
    text = text.lower()
    text = text.replace("\u2013", "-").replace("\u2014", "-")
    text = re.sub(r"[^\w\s\-\+\.#/]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def tokenize(text: str) -> list[str]:
    """Return normalized word tokens."""
    return [t for t in normalize_text(text).split() if len(t) > 1]


def get_search_variants(skill: str) -> list[str]:
    """Build normalized search phrases including aliases."""
    variants = {normalize_text(skill)}
    for alias in SKILL_ALIASES.get(skill.lower(), []):
        variants.add(normalize_text(alias))
    compact = skill.lower().replace(" ", "").replace("-", "").replace(".", "")
    if compact:
        variants.add(compact)
    return list(variants)


def _phrase_in_text(phrase: str, normalized_text: str, raw_lower: str) -> bool:
    """Check if a phrase appears with word boundaries or compact form."""
    if not phrase:
        return False
    if len(phrase) <= 3 and phrase.isalpha():
        pattern = rf"\b{re.escape(phrase)}\b"
        if re.search(pattern, normalized_text) or re.search(pattern, raw_lower):
            return True
    elif phrase in normalized_text or phrase in raw_lower:
        return True
    compact_phrase = phrase.replace(" ", "").replace("-", "").replace(".", "")
    compact_text = normalized_text.replace(" ", "")
    if compact_phrase and compact_phrase in compact_text:
        return True
    return False


def skill_in_resume(skill: str, raw_text: str, normalized_text: str) -> bool:
    """Smart skill detection: aliases, boundaries, partial/compact matching."""
    raw_lower = raw_text.lower()
    for variant in get_search_variants(skill):
        if _phrase_in_text(variant, normalized_text, raw_lower):
            return True
    skill_tokens = tokenize(skill)
    if len(skill_tokens) > 1:
        resume_tokens = set(tokenize(raw_text))
        if all(tok in resume_tokens for tok in skill_tokens):
            return True
    if len(skill) >= 4:
        for word in tokenize(raw_text):
            ratio = SequenceMatcher(None, normalize_text(skill), word).ratio()
            if ratio >= 0.88:
                return True
    return False


def match_skills(skills: list[str], raw_text: str) -> tuple[list[str], list[str]]:
    """Return matched and missing skills using improved matching."""
    if not skills:
        return [], []
    normalized = normalize_text(raw_text)
    matched, missing = [], []
    seen_matched = set()
    for skill in skills:
        if skill_in_resume(skill, raw_text, normalized):
            key = skill.lower()
            if key not in seen_matched:
                matched.append(skill)
                seen_matched.add(key)
        else:
            missing.append(skill)
    return matched, missing


def extract_meaningful_terms(text: str, min_len: int = 3) -> list[str]:
    """Extract candidate keywords from job description text."""
    normalized = normalize_text(text)
    tokens = [t for t in normalized.split() if t not in STOP_WORDS and len(t) >= min_len]
    bigrams = [f"{tokens[i]} {tokens[i+1]}" for i in range(len(tokens) - 1)]
    trigrams = [f"{tokens[i]} {tokens[i+1]} {tokens[i+2]}" for i in range(len(tokens) - 2)]
    counts = Counter(tokens + bigrams + trigrams)
    scored = []
    for term, freq in counts.items():
        if term in STOP_WORDS:
            continue
        weight = freq * (1.2 if " " in term else 1.0)
        scored.append((term, weight))
    scored.sort(key=lambda x: x[1], reverse=True)
    terms = []
    seen = set()
    for term, _ in scored:
        if term not in seen:
            terms.append(term)
            seen.add(term)
        if len(terms) >= 40:
            break
    return terms


def semantic_similarity(text_a: str, text_b: str) -> float:
    """Lightweight offline similarity using token overlap and sequence matching."""
    if not text_a.strip() or not text_b.strip():
        return 0.0
    tokens_a = set(tokenize(text_a)) - STOP_WORDS
    tokens_b = set(tokenize(text_b)) - STOP_WORDS
    if not tokens_a or not tokens_b:
        return 0.0
    intersection = tokens_a & tokens_b
    union = tokens_a | tokens_b
    jaccard = len(intersection) / len(union)
    seq_ratio = SequenceMatcher(None, normalize_text(text_a), normalize_text(text_b)).ratio()
    return round((jaccard * 0.65 + seq_ratio * 0.35) * 100, 1)


def match_job_description(resume_text: str, jd_text: str) -> dict:
    """Compare resume against optional job description."""
    jd_terms = extract_meaningful_terms(jd_text)
    matched, missing = match_skills(jd_terms, resume_text)
    semantic = semantic_similarity(resume_text, jd_text)
    keyword_score = (len(matched) / len(jd_terms) * 100) if jd_terms else 0
    match_pct = round(keyword_score * 0.7 + semantic * 0.3, 1)

    important_missing = missing[:8]
    recommendations = []
    if missing:
        recommendations.append(f"Highlight experience with: {', '.join(missing[:4])}.")
    if semantic < 45:
        recommendations.append("Mirror language from the job description in your summary and skills section.")
    if not re.search(r"\d+%|\d+\s*%|\$\d+|\d+\+", resume_text.lower()):
        recommendations.append("Quantify achievements with metrics (%, revenue, time saved, users impacted).")
    if "responsibilities" in normalize_text(jd_text) and "responsible" not in normalize_text(resume_text):
        recommendations.append("Align bullet points with key responsibilities mentioned in the job posting.")
    if len(matched) >= len(jd_terms) * 0.6:
        recommendations.append("Strong keyword overlap — tailor your summary to emphasize matched strengths.")
    if not recommendations:
        recommendations.append("Excellent alignment. Customize your cover letter to reference the company's mission.")

    return {
        "match_pct": match_pct,
        "semantic": semantic,
        "common": matched[:15],
        "missing": missing[:15],
        "important_missing": important_missing,
        "recommendations": recommendations[:6],
    }


# ---------------------------------------------------------------------------
# ATS scoring & analysis
# ---------------------------------------------------------------------------

def section_score(text: str, keywords: list[str]) -> float:
    normalized = normalize_text(text)
    hits = sum(1 for kw in keywords if kw in normalized)
    return min(hits / max(len(keywords) * 0.35, 1), 1.0)


def count_quantified_bullets(text: str) -> int:
    patterns = [
        r"\d+%", r"\$\d+", r"\d+\+", r"\d+\s*(years|yrs|months|users|clients|projects)",
        r"increased|decreased|reduced|improved|saved|grew|boosted",
    ]
    count = 0
    for line in text.splitlines():
        lower = line.lower()
        if any(re.search(p, lower) for p in patterns):
            count += 1
    return count


def formatting_score(text: str) -> float:
    score = 0.0
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    if len(lines) >= 8:
        score += 0.3
    bullet_lines = sum(1 for ln in lines if ln.startswith(("-", "•", "*", "·")))
    if bullet_lines >= 3:
        score += 0.3
    if len(text) >= 400:
        score += 0.2
    special_ratio = sum(1 for c in text if c in "{}[]|\\<>~`") / max(len(text), 1)
    if special_ratio < 0.01:
        score += 0.2
    return min(score, 1.0)


def calculate_ats_score(
    resume_text: str,
    role_skills: list[str],
    matched_skills: list[str],
) -> tuple[int, dict]:
    """Weighted ATS score across multiple resume dimensions."""
    normalized = normalize_text(resume_text)
    skill_ratio = len(matched_skills) / max(len(role_skills), 1)
    skills_pts = skill_ratio * 30

    exp_pts = section_score(resume_text, EXPERIENCE_KEYWORDS) * 20
    edu_pts = section_score(resume_text, EDUCATION_KEYWORDS) * 15
    proj_pts = section_score(resume_text, PROJECT_KEYWORDS) * 15

    completeness = 0
    if len(normalized) > 300:
        completeness += 0.35
    if len(normalized) > 800:
        completeness += 0.35
    sections_found = sum(
        1 for label in ["skills", "education", "experience", "project", "summary", "objective"]
        if label in normalized
    )
    completeness += min(sections_found * 0.05, 0.3)
    completeness_pts = min(completeness, 1.0) * 10

    fmt_pts = formatting_score(resume_text) * 10
    quant_bonus = min(count_quantified_bullets(resume_text) * 1.5, 5)

    total = int(round(skills_pts + exp_pts + edu_pts + proj_pts + completeness_pts + fmt_pts + quant_bonus))
    total = max(0, min(total, 100))

    breakdown = {
        "skills": round(skills_pts, 1),
        "experience": round(exp_pts, 1),
        "education": round(edu_pts, 1),
        "projects": round(proj_pts, 1),
        "completeness": round(completeness_pts, 1),
        "formatting": round(fmt_pts, 1),
        "quantified": round(quant_bonus, 1),
    }
    return total, breakdown


def score_status(score: int) -> tuple[str, str]:
    """Return human-readable status label and color."""
    if score >= 85:
        return "Excellent", SUCCESS
    if score >= 65:
        return "Good", WARNING
    if score >= 40:
        return "Average", ORANGE
    return "Needs Improvement", DANGER


def generate_strengths_weaknesses(matched: list[str], missing: list[str], resume_text: str) -> tuple[list[str], list[str]]:
    strengths = list(matched[:6])
    if count_quantified_bullets(resume_text) >= 2 and "Quantified achievements" not in strengths:
        strengths.append("Quantified achievements")
    if section_score(resume_text, PROJECT_KEYWORDS) > 0.4 and "Project portfolio" not in strengths:
        strengths.append("Project portfolio")
    if section_score(resume_text, EDUCATION_KEYWORDS) > 0.3 and "Education credentials" not in strengths:
        strengths.append("Education credentials")

    weaknesses = [f"Missing {m}" for m in missing[:5]]
    if count_quantified_bullets(resume_text) < 2:
        weaknesses.append("Lacks measurable impact metrics")
    if section_score(resume_text, PROJECT_KEYWORDS) < 0.25:
        weaknesses.append("Limited project highlights")
    if section_score(resume_text, EXPERIENCE_KEYWORDS) < 0.25:
        weaknesses.append("Experience section needs detail")
    return strengths[:6], weaknesses[:6]


def generate_ai_suggestions(
    role: str,
    matched: list[str],
    missing: list[str],
    resume_text: str,
    jd_result: dict | None,
) -> list[str]:
    suggestions = []
    normalized = normalize_text(resume_text)

    for skill in missing[:4]:
        suggestions.append(f"Add {skill} experience or coursework to strengthen role alignment.")

    role_certs = CERT_SUGGESTIONS.get(role, [])
    for cert in role_certs[:2]:
        suggestions.append(f"Consider certification: {cert}.")

    if "github" not in normalized and "gitlab" not in normalized:
        suggestions.append("Mention GitHub projects with links to demonstrate practical work.")
    if "docker" not in normalized and role in ("Python Developer", "Cloud Engineer", "AI / ML Engineer"):
        suggestions.append("Add Docker experience to show deployment and containerization skills.")
    if not re.search(r"\d+", resume_text):
        suggestions.append("Quantify achievements with numbers (team size, performance gains, revenue, users).")
    if section_score(resume_text, PROJECT_KEYWORDS) < 0.3:
        suggestions.append("Include 2–3 project entries with tech stack and measurable outcomes.")
    if "led" not in normalized and "managed" not in normalized:
        suggestions.append("Use stronger action verbs: Led, Designed, Implemented, Optimized, Delivered.")

    tech_gaps = [m for m in missing if m in (
        "aws", "azure", "docker", "kubernetes", "react", "sql", "rest", "api", "tensorflow", "pytorch",
    )]
    for tech in tech_gaps[:3]:
        suggestions.append(f"Include hands-on experience with {tech} or related tools.")

    if jd_result and jd_result.get("important_missing"):
        jd_miss = jd_result["important_missing"][:3]
        suggestions.append(f"Job description gap — add keywords: {', '.join(jd_miss)}.")

    seen = set()
    unique = []
    for s in suggestions:
        if s not in seen:
            unique.append(s)
            seen.add(s)
    return unique[:10]


def generate_overall_impression(role: str, ats_score: int, matched: list[str], missing: list[str]) -> str:
    if ats_score >= 85:
        tone = "Excellent candidate profile"
    elif ats_score >= 65:
        tone = "Good candidate"
    elif ats_score >= 40:
        tone = "Moderate fit"
    else:
        tone = "Needs significant improvement"

    if matched and missing:
        return (
            f"{tone} for {role} with strengths in {', '.join(matched[:3])}. "
            f"Improve visibility by addressing {', '.join(missing[:3])}."
        )
    if matched:
        return f"{tone} for {role}. Strong coverage of core skills: {', '.join(matched[:4])}."
    return f"{tone} for {role}. Add role-specific skills and project evidence to improve ATS performance."


def build_analysis_report(
    role: str,
    resume_text: str,
    jd_text: str,
    ml_classifier: RoleClassifier | None = None,
) -> dict:
    """Run full analysis pipeline and return structured results."""
    skills = JOB_KEYWORDS[role]
    matched, missing = match_skills(skills, resume_text)
    ats_score, breakdown = calculate_ats_score(resume_text, skills, matched)
    status, _ = score_status(ats_score)
    strengths, weaknesses = generate_strengths_weaknesses(matched, missing, resume_text)

    jd_result = None
    if jd_text.strip() and jd_text.strip() != PLACEHOLDER_JD:
        jd_result = match_job_description(resume_text, jd_text)

    suggestions = generate_ai_suggestions(role, matched, missing, resume_text, jd_result)
    impression = generate_overall_impression(role, ats_score, matched, missing)

    # Supervised learning prediction — additional intelligent layer on top of keyword matching
    ml_prediction = None
    if ml_classifier and ml_classifier.is_trained:
        ml_prediction = ml_classifier.predict(resume_text)

    tips = [
        "Keep formatting ATS-friendly: simple headings, no tables or text boxes.",
        "Place a concise skills section near the top for recruiter scanning.",
        "Tailor your resume summary to the target role in 2–3 lines.",
        "Use consistent date formats and reverse-chronological work history.",
    ]
    if count_quantified_bullets(resume_text) < 2:
        tips.insert(0, "Add metrics to every major bullet point (%, $, time saved, scale).")

    return {
        "role": role,
        "ats_score": ats_score,
        "ats_breakdown": breakdown,
        "status": status,
        "matched": matched,
        "missing": missing,
        "strengths": strengths,
        "weaknesses": weaknesses,
        "suggestions": suggestions,
        "tips": tips,
        "impression": impression,
        "jd_result": jd_result,
        "ml_prediction": ml_prediction,
    }


def format_report(data: dict) -> str:
    """Format analysis dict into a readable text report."""
    lines = [
        "═" * 52,
        "  AI SMART RESUME ANALYSIS REPORT",
        "═" * 52,
        "",
        f"Selected Role:     {data['role']}",
    ]
    if data.get("ml_prediction"):
        ml = data["ml_prediction"]
        lines.append(f"ML Predicted Role: {ml['predicted_role']}")
        lines.append(f"ML Confidence:     {ml['confidence']}%")
    lines.append(f"ATS Score:         {data['ats_score']} / 100")
    if data["jd_result"]:
        lines.append(f"Resume Match:      {data['jd_result']['match_pct']}%")
        lines.append(f"Semantic Similarity: {data['jd_result']['semantic']}%")
    lines.extend([
        f"Status:            {data['status']}",
        "",
        "─" * 52,
        "SUPERVISED ML PREDICTION",
        "─" * 52,
    ])
    if data.get("ml_prediction"):
        ml = data["ml_prediction"]
        lines.extend([
            f"  Predicted Role:    {ml['predicted_role']}",
            f"  Confidence:        {ml['confidence']}%",
            f"  Model:             Multinomial Naive Bayes + TF-IDF",
            f"  Note:              ML prediction complements ATS keyword scoring.",
        ])
    else:
        lines.append("  ML prediction unavailable.")

    lines.extend([
        "",
        "─" * 52,
        "ATS BREAKDOWN",
        "─" * 52,
        f"  Skills:        {data['ats_breakdown']['skills']:.1f} / 30",
        f"  Experience:    {data['ats_breakdown']['experience']:.1f} / 20",
        f"  Education:     {data['ats_breakdown']['education']:.1f} / 15",
        f"  Projects:      {data['ats_breakdown']['projects']:.1f} / 15",
        f"  Completeness:  {data['ats_breakdown']['completeness']:.1f} / 10",
        f"  Formatting:    {data['ats_breakdown']['formatting']:.1f} / 10",
        f"  Quantified:    {data['ats_breakdown']['quantified']:.1f} / 5 bonus",
        "",
        "─" * 52,
        "MATCHED SKILLS",
        "─" * 52,
        "  " + (", ".join(data["matched"]) if data["matched"] else "None detected"),
        "",
        "─" * 52,
        "MISSING SKILLS",
        "─" * 52,
        "  " + (", ".join(data["missing"]) if data["missing"] else "None — great coverage!"),
    ])

    if data["jd_result"]:
        jd = data["jd_result"]
        lines.extend([
            "",
            "─" * 52,
            "JOB DESCRIPTION MATCH",
            "─" * 52,
            f"  Common Skills:   {', '.join(jd['common']) if jd['common'] else 'None'}",
            f"  Missing Skills:  {', '.join(jd['missing'][:10]) if jd['missing'] else 'None'}",
            f"  Critical Gaps:   {', '.join(jd['important_missing']) if jd['important_missing'] else 'None'}",
            "",
            "  Recommendations:",
        ])
        for rec in jd["recommendations"]:
            lines.append(f"    • {rec}")

    lines.extend([
        "",
        "─" * 52,
        "STRENGTHS",
        "─" * 52,
    ])
    for s in data["strengths"]:
        lines.append(f"  • {s}")

    lines.extend(["", "─" * 52, "WEAKNESSES", "─" * 52])
    for w in data["weaknesses"]:
        lines.append(f"  • {w}")

    lines.extend(["", "─" * 52, "AI SUGGESTIONS", "─" * 52])
    for s in data["suggestions"]:
        lines.append(f"  • {s}")

    lines.extend(["", "─" * 52, "RESUME IMPROVEMENT TIPS", "─" * 52])
    for t in data["tips"]:
        lines.append(f"  • {t}")

    lines.extend([
        "",
        "─" * 52,
        "OVERALL IMPRESSION",
        "─" * 52,
        f"  {data['impression']}",
        "",
        "═" * 52,
    ])
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# File loading
# ---------------------------------------------------------------------------

def extract_text_from_file(path: str) -> str:
    """Extract plain text from TXT, PDF, or DOCX files."""
    lower = path.lower()
    if lower.endswith(".txt"):
        with open(path, encoding="utf-8", errors="ignore") as f:
            return f.read()
    if lower.endswith(".pdf"):
        if pdfplumber is None:
            raise ImportError("pdfplumber is not installed. Run: pip install pdfplumber")
        chunks = []
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    chunks.append(text)
        return "\n".join(chunks)
    if lower.endswith(".docx"):
        if Document is None:
            raise ImportError("python-docx is not installed. Run: pip install python-docx")
        doc = Document(path)
        return "\n".join(p.text for p in doc.paragraphs if p.text.strip())
    raise ValueError("Unsupported file type. Please use .txt, .pdf, or .docx")


def is_placeholder(text: str, placeholder: str) -> bool:
    return not text.strip() or text.strip() == placeholder


# ---------------------------------------------------------------------------
# Application UI
# ---------------------------------------------------------------------------

class ResumeAnalyzerApp(ctk.CTk):
    """Main application window."""

    def __init__(self, ml_classifier: RoleClassifier | None = None):
        super().__init__()
        self.ml_classifier = ml_classifier or ROLE_CLASSIFIER
        self.title("AI Smart Resume Analyzer")
        self.geometry("1380x860")
        self.minsize(1100, 700)
        self.configure(fg_color=BG)
        self._build_ui()

    def _build_ui(self):
        self._build_header()
        self._build_main_layout()
        self._build_footer()

    def _build_header(self):
        header_frame = ctk.CTkFrame(self, fg_color=SURFACE, corner_radius=0, height=100)
        header_frame.pack(fill="x")
        header_frame.pack_propagate(False)

        inner = ctk.CTkFrame(header_frame, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=28, pady=18)

        title = ctk.CTkLabel(
            inner,
            text="AI Smart Resume Analyzer",
            font=(FONT_FAMILY, 32, "bold"),
            text_color=TEXT_PRIMARY,
        )
        title.pack(anchor="w")

        subtitle = ctk.CTkLabel(
            inner,
            text="Offline ATS scoring • Supervised ML role prediction • Job description matching",
            font=(FONT_FAMILY, 14),
            text_color=TEXT_SECONDARY,
        )
        subtitle.pack(anchor="w", pady=(4, 0))

        badge = ctk.CTkLabel(
            inner,
            text="  ●  100% Local Processing  ",
            font=(FONT_FAMILY, 12, "bold"),
            fg_color=ACCENT,
            corner_radius=20,
            text_color=TEXT_PRIMARY,
        )
        badge.place(relx=1.0, rely=0.5, anchor="e")

    def _card(self, parent, **kwargs) -> ctk.CTkFrame:
        return ctk.CTkFrame(
            parent,
            fg_color=CARD,
            corner_radius=16,
            border_width=1,
            border_color=BORDER,
            **kwargs,
        )

    def _section_label(self, parent, text: str, size: int = 15):
        return ctk.CTkLabel(
            parent,
            text=text,
            font=(FONT_FAMILY, size, "bold"),
            text_color=TEXT_PRIMARY,
            anchor="w",
        )

    def _build_main_layout(self):
        container = ctk.CTkFrame(self, fg_color=BG)
        container.pack(fill="both", expand=True, padx=24, pady=(16, 8))
        container.grid_columnconfigure(0, weight=3)
        container.grid_columnconfigure(1, weight=2)
        container.grid_rowconfigure(0, weight=1)

        # --- Left panel (inputs) ---
        left = self._card(container)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left.grid_rowconfigure(3, weight=1)
        left.grid_rowconfigure(5, weight=1)
        left.grid_columnconfigure(0, weight=1)

        self._section_label(left, "Target Role").grid(row=0, column=0, sticky="w", padx=20, pady=(20, 6))
        self.role_var = ctk.StringVar(value=list(JOB_KEYWORDS.keys())[0])
        self.role_menu = ctk.CTkOptionMenu(
            left,
            variable=self.role_var,
            values=list(JOB_KEYWORDS.keys()),
            width=320,
            height=38,
            fg_color=SURFACE,
            button_color=ACCENT,
            button_hover_color=ACCENT_LIGHT,
            dropdown_fg_color=CARD,
            font=(FONT_FAMILY, 14),
        )
        self.role_menu.grid(row=1, column=0, sticky="ew", padx=20)

        self._section_label(left, "Resume").grid(row=2, column=0, sticky="w", padx=20, pady=(16, 6))
        self.resume_box = ctk.CTkTextbox(
            left,
            height=240,
            corner_radius=12,
            fg_color=SURFACE,
            border_width=1,
            border_color=BORDER,
            font=(FONT_FAMILY, 13),
            wrap="word",
        )
        self.resume_box.grid(row=3, column=0, sticky="nsew", padx=20, pady=(0, 8))
        self.resume_box.insert("1.0", PLACEHOLDER_RESUME)

        self._section_label(left, "Job Description (Optional)").grid(row=4, column=0, sticky="w", padx=20, pady=(8, 6))
        self.jd_box = ctk.CTkTextbox(
            left,
            height=120,
            corner_radius=12,
            fg_color=SURFACE,
            border_width=1,
            border_color=BORDER,
            font=(FONT_FAMILY, 13),
            wrap="word",
        )
        self.jd_box.grid(row=5, column=0, sticky="nsew", padx=20, pady=(0, 12))
        self.jd_box.insert("1.0", PLACEHOLDER_JD)

        btn_row = ctk.CTkFrame(left, fg_color="transparent")
        btn_row.grid(row=6, column=0, sticky="ew", padx=20, pady=(0, 20))
        for i in range(5):
            btn_row.grid_columnconfigure(i, weight=1)

        self._mk_btn(btn_row, "Analyze Resume", self.analyze_resume, ACCENT, 0)
        self._mk_btn(btn_row, "Load Resume", self.load_resume, SURFACE, 1)
        self._mk_btn(btn_row, "Save Report", self.save_report, SURFACE, 2)
        self._mk_btn(btn_row, "Clear", self.clear_all, "#374151", 3)
        self._mk_btn(btn_row, "Exit", self.quit_app, DANGER, 4)

        # --- Right panel (dashboard) ---
        right = self._card(container)
        right.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        right.grid_columnconfigure(0, weight=1)

        self._section_label(right, "Analysis Dashboard", 16).grid(row=0, column=0, sticky="w", padx=20, pady=(20, 12))

        metrics = ctk.CTkFrame(right, fg_color="transparent")
        metrics.grid(row=1, column=0, sticky="ew", padx=20)
        metrics.grid_columnconfigure((0, 1, 2), weight=1)

        self.ats_card = self._metric_card(metrics, "ATS Score", "0", "/ 100", 0)
        self.ml_card = self._metric_card(metrics, "ML Confidence", "—", "", 1)
        self.match_card = self._metric_card(metrics, "Job Match", "—", "", 2)

        self.ml_role_label = ctk.CTkLabel(
            right,
            text="ML Predicted Role: —",
            font=(FONT_FAMILY, 14, "bold"),
            text_color=ACCENT_LIGHT,
        )
        self.ml_role_label.grid(row=2, column=0, sticky="w", padx=20, pady=(10, 0))

        self.status_label = ctk.CTkLabel(
            right,
            text="Status: Awaiting Analysis",
            font=(FONT_FAMILY, 15, "bold"),
            text_color=TEXT_SECONDARY,
        )
        self.status_label.grid(row=3, column=0, sticky="w", padx=20, pady=(8, 6))

        prog_frame = ctk.CTkFrame(right, fg_color="transparent")
        prog_frame.grid(row=4, column=0, sticky="ew", padx=20, pady=(0, 10))
        prog_frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(prog_frame, text="ATS Progress", font=(FONT_FAMILY, 12), text_color=TEXT_MUTED).grid(
            row=0, column=0, sticky="w"
        )
        self.ats_progress = ctk.CTkProgressBar(
            prog_frame, height=14, corner_radius=8, progress_color=ACCENT, fg_color=SURFACE
        )
        self.ats_progress.grid(row=1, column=0, sticky="ew", pady=(4, 8))
        self.ats_progress.set(0)

        ctk.CTkLabel(prog_frame, text="Job Match Progress", font=(FONT_FAMILY, 12), text_color=TEXT_MUTED).grid(
            row=2, column=0, sticky="w"
        )
        self.match_progress = ctk.CTkProgressBar(
            prog_frame, height=14, corner_radius=8, progress_color=SUCCESS, fg_color=SURFACE
        )
        self.match_progress.grid(row=3, column=0, sticky="ew", pady=(4, 0))
        self.match_progress.set(0)

        self._section_label(right, "Detailed Report").grid(row=5, column=0, sticky="nw", padx=20, pady=(8, 6))
        self.result_box = ctk.CTkTextbox(
            right,
            corner_radius=12,
            fg_color=SURFACE,
            border_width=1,
            border_color=BORDER,
            font=(FONT_FAMILY, 12),
            wrap="word",
        )
        self.result_box.grid(row=6, column=0, sticky="nsew", padx=20, pady=(0, 20))
        right.grid_rowconfigure(6, weight=1)
        self.result_box.insert("1.0", "Run analysis to generate your personalized resume report.")

    def _metric_card(self, parent, title: str, value: str, suffix: str, col: int) -> dict:
        card = ctk.CTkFrame(parent, fg_color=SURFACE, corner_radius=14, border_width=1, border_color=BORDER)
        card.grid(row=0, column=col, sticky="ew", padx=(0 if col == 0 else 6, 6 if col == 0 else 0), pady=4)
        ctk.CTkLabel(card, text=title, font=(FONT_FAMILY, 12), text_color=TEXT_MUTED).pack(anchor="w", padx=14, pady=(12, 0))
        row = ctk.CTkFrame(card, fg_color="transparent")
        row.pack(anchor="w", padx=14, pady=(4, 12))
        val_lbl = ctk.CTkLabel(row, text=value, font=(FONT_FAMILY, 28, "bold"), text_color=TEXT_PRIMARY)
        val_lbl.pack(side="left")
        suf_lbl = ctk.CTkLabel(row, text=suffix, font=(FONT_FAMILY, 14), text_color=TEXT_SECONDARY)
        suf_lbl.pack(side="left", padx=(4, 0), pady=(8, 0))
        return {"value": val_lbl, "suffix": suf_lbl}

    def _mk_btn(self, parent, text, command, color, col):
        ctk.CTkButton(
            parent,
            text=text,
            command=command,
            height=38,
            fg_color=color,
            hover_color=ACCENT_LIGHT if color == ACCENT else CARD_HOVER,
            font=(FONT_FAMILY, 13, "bold"),
            corner_radius=10,
        ).grid(row=0, column=col, sticky="ew", padx=4)

    def _build_footer(self):
        footer = ctk.CTkLabel(
            self,
            text="Supports TXT • PDF • DOCX  |  Scikit-learn ML  |  No internet required",
            font=(FONT_FAMILY, 11),
            text_color=TEXT_MUTED,
        )
        footer.pack(pady=(0, 10))

    def _get_resume_text(self) -> str | None:
        text = self.resume_box.get("1.0", "end")
        if is_placeholder(text, PLACEHOLDER_RESUME):
            messagebox.showwarning("Missing Resume", "Please paste or load your resume before analyzing.")
            return None
        return text

    def _get_jd_text(self) -> str:
        text = self.jd_box.get("1.0", "end")
        if is_placeholder(text, PLACEHOLDER_JD):
            return ""
        return text

    def analyze_resume(self):
        resume_text = self._get_resume_text()
        if resume_text is None:
            return

        role = self.role_var.get()
        jd_text = self._get_jd_text()
        data = build_analysis_report(role, resume_text, jd_text, self.ml_classifier)

        ats = data["ats_score"]
        status, color = score_status(ats)
        self.ats_card["value"].configure(text=str(ats))
        self.ats_progress.set(ats / 100)

        if data.get("ml_prediction"):
            ml = data["ml_prediction"]
            self.ml_card["value"].configure(text=str(ml["confidence"]))
            self.ml_card["suffix"].configure(text="%")
            self.ml_role_label.configure(text=f"ML Predicted Role: {ml['predicted_role']}")
        else:
            self.ml_card["value"].configure(text="—")
            self.ml_card["suffix"].configure(text="")
            self.ml_role_label.configure(text="ML Predicted Role: —")

        if data["jd_result"]:
            match_pct = data["jd_result"]["match_pct"]
            self.match_card["value"].configure(text=str(match_pct))
            self.match_card["suffix"].configure(text="%")
            self.match_progress.set(match_pct / 100)
        else:
            self.match_card["value"].configure(text="—")
            self.match_card["suffix"].configure(text="")
            self.match_progress.set(0)

        self.status_label.configure(text=f"Status: {status}", text_color=color)
        self.result_box.delete("1.0", "end")
        self.result_box.insert("1.0", format_report(data))

    def load_resume(self):
        path = filedialog.askopenfilename(
            title="Load Resume",
            filetypes=[
                ("Supported Files", "*.txt *.pdf *.docx"),
                ("Text Files", "*.txt"),
                ("PDF Files", "*.pdf"),
                ("Word Documents", "*.docx"),
                ("All Files", "*.*"),
            ],
        )
        if not path:
            return
        try:
            content = extract_text_from_file(path)
            if not content.strip():
                messagebox.showwarning("Empty File", "No text could be extracted from the selected file.")
                return
            self.resume_box.delete("1.0", "end")
            self.resume_box.insert("1.0", content)
        except ImportError as exc:
            messagebox.showerror("Missing Dependency", str(exc))
        except Exception as exc:
            messagebox.showerror("Load Error", f"Could not read file:\n{exc}")

    def save_report(self):
        report = self.result_box.get("1.0", "end").strip()
        if not report or report.startswith("Run analysis"):
            messagebox.showinfo("Nothing to Save", "Analyze a resume first to generate a report.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            initialfile=f"resume_report_{datetime.now().strftime('%Y%m%d_%H%M')}.txt",
            filetypes=[("Text Report", "*.txt"), ("All Files", "*.*")],
        )
        if path:
            with open(path, "w", encoding="utf-8") as f:
                f.write(report)
            messagebox.showinfo("Saved", f"Report saved to:\n{path}")

    def clear_all(self):
        self.resume_box.delete("1.0", "end")
        self.resume_box.insert("1.0", PLACEHOLDER_RESUME)
        self.jd_box.delete("1.0", "end")
        self.jd_box.insert("1.0", PLACEHOLDER_JD)
        self.result_box.delete("1.0", "end")
        self.result_box.insert("1.0", "Run analysis to generate your personalized resume report.")
        self.ats_card["value"].configure(text="0")
        self.ml_card["value"].configure(text="—")
        self.ml_card["suffix"].configure(text="")
        self.ml_role_label.configure(text="ML Predicted Role: —")
        self.match_card["value"].configure(text="—")
        self.match_card["suffix"].configure(text="")
        self.ats_progress.set(0)
        self.match_progress.set(0)
        self.status_label.configure(text="Status: Awaiting Analysis", text_color=TEXT_SECONDARY)

    def quit_app(self):
        self.destroy()


def main():
    # Step: Train supervised learning model automatically at application startup
    classifier = train_ml_model_at_startup()
    print(f"[ML] Model trained on {len(ML_TRAINING_DATA)} labeled samples.")
    print(f"[ML] Training accuracy: {classifier.training_accuracy}%")

    app = ResumeAnalyzerApp(ml_classifier=classifier)
    app.mainloop()


if __name__ == "__main__":
    main()
