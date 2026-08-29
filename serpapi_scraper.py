#!/usr/bin/env python3
"""
serpapi_scraper.py

Script minimal pour:
- interroger SerpApi (engine=google) avec des dorks (par défaut: LinkedIn et CV)
- récupérer les pages résultats
- extraire par heuristique: nom, ville, expériences, études, compétences, certifications, années d'expérience par compétence, url
- sauvegarder les résultats dans un fichier Excel

Usage:
  export SERPAPI_API_KEY="<votre_cle>"
  python serpapi_scraper.py --input input_names.csv --output results.xlsx

Ou passer la clé en CLI: --api-key <KEY>

Entrée attendue (input_names.csv) : CSV avec colonnes (prenom,nom,ville) — ville optionnelle

Dépendances:
  pip install requests beautifulsoup4 pandas openpyxl

Remarques:
  - Le scraping est heuristique; pour des sites spécifiques (LinkedIn, CV template), adapter les sélecteurs.
  - Respecter les conditions d'utilisation des sites et la loi.
"""

import os
import time
import re
import argparse
import csv
import json
from typing import List, Dict, Optional

import pandas as pd

import requests
from bs4 import BeautifulSoup

SERPAPI_SEARCH_URL = 'https://serpapi.com/search.json'

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (compatible; SerpApiScraper/1.0; +https://example.com)'
}

# Heuristiques pour repérer sections
SECTION_KEYWORDS = {
    'experience': ['experience', 'expérience', 'expériences', 'professional experience', 'work experience'],
    'education': ['education', 'formation', 'éducation', 'études', 'formations'],
    'skills': ['skills', 'compétences', 'competences', 'technologies'],
    'certifications': ['certification', 'certifications', 'certifié', 'certifie']
}

SKILL_YEAR_PATTERNS = [
    re.compile(r"([A-Za-z+#\. ]{2,40})\s*[\-\(\[]\s*(\d+)\s*(?:ans|years)\b", re.I),  # Python (5 ans)
    re.compile(r"(\d+)\s*(?:ans|years)\s+of\s+([A-Za-z+#\. ]{2,40})", re.I),  # 5 years of Python
]

# Normalized technology keywords (français + english) mapping canonical name -> variants to search for
TECHNOLOGY_KEYWORDS = {
    'Python': ['python', 'python3', 'pyhton'],
    'Java': ['java'],
    'JavaScript': ['javascript', 'js'],
    'TypeScript': ['typescript'],
    'SQL': ['sql'],
    'PostgreSQL': ['postgres', 'postgresql'],
    'MySQL': ['mysql'],
    'MongoDB': ['mongodb', 'mongo db'],
    'Spark': ['spark', 'apache spark'],
    'Hadoop': ['hadoop'],
    'Kafka': ['kafka', 'apache kafka'],
    'Airflow': ['airflow', 'apache airflow'],
    'Databricks': ['databricks'],
    'AWS': ['aws', 'amazon web services', 'amazon webservice', 'amazon'],
    'Azure': ['azure', 'microsoft azure'],
    'GCP': ['gcp', 'google cloud', 'google cloud platform'],
    'Docker': ['docker', 'containers'],
    'Kubernetes': ['kubernetes', 'k8s'],
    'Terraform': ['terraform'],
    'Power BI': ['power bi', 'powerbi'],
    'Tableau': ['tableau'],
    'SAP Datasphere': ['sap datasphere', 'datasphere'],
    'SAP S/4HANA': ['s/4hana', 's4hana', 'sap s/4hana'],
    'Databricks': ['databricks'],
    'Pandas': ['pandas'],
    'NumPy': ['numpy', 'numPy'],
    'scikit-learn': ['scikit-learn', 'sklearn'],
    'TensorFlow': ['tensorflow'],
    'PyTorch': ['pytorch'],
    'Machine Learning': ['machine learning', 'apprentissage automatique', 'apprentissage profond', 'deep learning', 'ml'],
    'ETL': ['etl', 'extraction transformation chargement', 'extract transform load'],
    'BigQuery': ['bigquery'],
    'PowerShell': ['powershell'],
    'Shell': ['bash', 'sh', 'shell'],
}

# Precompile regex patterns for faster detection
_TECH_PATTERNS = []
for canon, variants in TECHNOLOGY_KEYWORDS.items():
    # join variants into regex alternation, escape special chars
    alt = '|'.join(re.escape(v) for v in variants)
    # word boundaries but allow plus/ # in names
    pattern = re.compile(r'(?<!\w)(' + alt + r')(?!\w)', re.I)
    _TECH_PATTERNS.append((canon, pattern))


def detect_technologies(text: str) -> List[str]:
    """Detect known technologies in the provided text and return canonical names (deduplicated).
    Supports French and English variants defined in TECHNOLOGY_KEYWORDS."""
    if not text:
        return []
    found = []
    for canon, pat in _TECH_PATTERNS:
        if pat.search(text):
            found.append(canon)
    # preserve original order defined by TECHNOLOGY_KEYWORDS
    uniq = []
    seen = set()
    for canon in TECHNOLOGY_KEYWORDS.keys():
        if canon in found and canon not in seen:
            uniq.append(canon)
            seen.add(canon)
    return uniq

# Map canonical technology names to categories (French labels)
TECHNOLOGY_CATEGORIES = {
    'Python': 'Langages',
    'Java': 'Langages',
    'JavaScript': 'Langages',
    'TypeScript': 'Langages',
    'SQL': 'Langages',
    'PostgreSQL': 'Bases de données',
    'MySQL': 'Bases de données',
    'MongoDB': 'Bases de données',
    'Spark': 'Big Data / Traitement',
    'Hadoop': 'Big Data / Traitement',
    'Kafka': 'Streaming / Messaging',
    'Airflow': 'Orchestration',
    'Databricks': 'Plateformes Data',
    'AWS': 'Cloud',
    'Azure': 'Cloud',
    'GCP': 'Cloud',
    'Docker': 'Conteneurs',
    'Kubernetes': 'Orchestration',
    'Terraform': 'Infra as Code',
    'Power BI': 'BI',
    'Tableau': 'BI',
    'SAP Datasphere': 'Plateformes Data',
    'SAP S/4HANA': 'ERP',
    'Pandas': 'Librairies Data',
    'NumPy': 'Librairies Data',
    'scikit-learn': 'ML / IA',
    'TensorFlow': 'ML / IA',
    'PyTorch': 'ML / IA',
    'Machine Learning': 'ML / IA',
    'ETL': 'ETL',
    'BigQuery': 'Bases de données',
    'PowerShell': 'Scripting',
    'Shell': 'Scripting',
}


def categorize_technologies(tech_list: List[str]) -> str:
    """Return a semicolon-separated string grouping technologies by category.
    Example: 'Langages: Python, Java; Cloud: AWS, Azure' (French category labels).
    """
    if not tech_list:
        return ''
    buckets = {}
    for tech in tech_list:
        cat = TECHNOLOGY_CATEGORIES.get(tech, 'Autre')
        buckets.setdefault(cat, []).append(tech)
    parts = []
    # define preferred order for categories
    preferred = ['Langages', 'Librairies Data', 'ML / IA', 'Big Data / Traitement', 'Plateformes Data', 'Bases de données', 'Cloud', 'Conteneurs', 'Orchestration', 'Infra as Code', 'BI', 'Streaming / Messaging', 'ETL', 'Scripting', 'ERP', 'Autre']
    for cat in preferred:
        vals = buckets.get(cat)
        if vals:
            parts.append(f"{cat}: {', '.join(vals)}")
    # include any other categories not in preferred
    for cat, vals in buckets.items():
        if cat not in preferred:
            parts.append(f"{cat}: {', '.join(vals)}")
    return ' ; '.join(parts)



def search_serpapi(query: str, api_key: str, num: int = 10) -> List[Dict]:
    params = {
        'engine': 'google',
        'q': query,
        'num': num,
        'api_key': api_key,
    }
    resp = requests.get(SERPAPI_SEARCH_URL, params=params, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    results = []
    # SerpApi places organic results in 'organic_results'
    for item in data.get('organic_results', []):
        link = item.get('link') or item.get('snippet')
        results.append({
            'title': item.get('title'),
            'link': link,
            'snippet': item.get('snippet'),
            'rich_snippet': item.get('rich_snippet'),
            'raw': item
        })
    return results


def fetch_page(url: str, timeout: int = 20) -> Optional[str]:
    try:
        r = requests.get(url, headers=HEADERS, timeout=timeout)
        r.raise_for_status()
        return r.text
    except Exception as e:
        print(f"[fetch_page] Erreur pour {url}: {e}")
        return None


def text_of_tag(tag):
    if not tag:
        return ''
    return ' '.join(tag.stripped_strings)


def find_section_text(soup: BeautifulSoup, keywords: List[str]) -> str:
    # Cherche titres qui correspondent et récupère le texte adjacent (ul/li ou paragraphes)
    for header in soup.find_all(['h1', 'h2', 'h3', 'h4', 'strong', 'b']):
        htext = header.get_text(separator=' ').strip().lower()
        for kw in keywords:
            if kw in htext:
                # cherche une liste immédiate ou paragraphes suivants
                parent = header.parent
                # case: ul/ol following
                ul = header.find_next(['ul', 'ol'])
                if ul:
                    items = [li.get_text(separator=' ').strip() for li in ul.find_all('li')]
                    if items:
                        return '\n'.join(items)
                # else collect following paragraphs until next header
                texts = []
                for sib in header.next_siblings:
                    if getattr(sib, 'name', None) and sib.name in ['h1', 'h2', 'h3', 'h4']:
                        break
                    if getattr(sib, 'get_text', None):
                        t = sib.get_text(separator=' ').strip()
                        if t:
                            texts.append(t)
                if texts:
                    return '\n'.join(texts)
                # fallback: parent text
                ptext = parent.get_text(separator=' ').strip()
                if ptext:
                    return ptext
    # fallback: search by keyword anywhere in page and return nearby text
    body = soup.get_text(separator='\n')
    for kw in keywords:
        idx = body.lower().find(kw)
        if idx != -1:
            start = max(0, idx - 400)
            end = min(len(body), idx + 800)
            return body[start:end].strip()
    return ''


def extract_name_city(soup: BeautifulSoup) -> (str, str):
    # Nom: souvent en h1 ou dans le title
    name = ''
    if soup.h1:
        name = soup.h1.get_text(separator=' ').strip()
    if not name and soup.title:
        title = soup.title.get_text(separator=' ')
        # nettoie si title contient "| LinkedIn"
        name = re.sub(r"\s*\|.*$", '', title).strip()
    # Ville heuristique: chercher mot-clés "ville" ou pattern 'Location' ou texte court après 'Lives in' ou 'Basé à'
    city = ''
    text = soup.get_text(separator='\n')
    m = re.search(r"(Lives in|Basé à|Basée à|Based in)[:\s]+([A-Za-z\-\s,]+)", text, re.I)
    if m:
        city = m.group(2).split('\n')[0].strip()
    else:
        # chercher petite ligne après 'Location' ou 'Ville'
        m2 = re.search(r"(Location|Ville)[:\s]*\n?\s*([A-Za-z\-\s,]+)", text, re.I)
        if m2:
            city = m2.group(2).strip()
    # limit length
    if city and len(city) > 100:
        city = city[:100]
    return name, city


def extract_skills_and_years(skills_text: str) -> (List[str], Dict[str, int]):
    skills = []
    skill_years = {}
    if not skills_text:
        return skills, skill_years
    # split lines and commas
    parts = re.split(r"[\n,;]+", skills_text)
    for p in parts:
        p = p.strip()
        if not p:
            continue
        # try to extract "Skill (5 ans)" patterns
        matched = False
        for pat in SKILL_YEAR_PATTERNS:
            m = pat.search(p)
            if m:
                # pattern groups may vary
                if len(m.groups()) >= 2:
                    g1, g2 = m.group(1).strip(), m.group(2).strip()
                    # depending on pattern order, ensure skill name and years
                    if g1.isdigit():
                        years = int(g1)
                        skill = m.group(2).strip()
                    else:
                        skill = g1
                        years = int(g2)
                    skills.append(skill)
                    skill_years[skill] = years
                    matched = True
                    break
        if not matched:
            # remove any trailing years like "- 5 ans" or "— 5 years"
            m2 = re.search(r"([A-Za-z0-9+#+\.\s]{1,60})[\-–—]\s*(\d+)\s*(?:ans|years)", p, re.I)
            if m2:
                skill = m2.group(1).strip()
                years = int(m2.group(2))
                skills.append(skill)
                skill_years[skill] = years
                matched = True
        if not matched:
            # fallback: treat p as skill list
            skills.append(p)
    # deduplicate preserving order
    seen = set()
    uniq_skills = []
    for s in skills:
        key = s.lower()
        if key not in seen:
            seen.add(key)
            uniq_skills.append(s)
    return uniq_skills, skill_years


def parse_profile(html: str, url: str) -> Dict:
    soup = BeautifulSoup(html, 'html.parser')
    title = soup.title.get_text(separator=' ').strip() if soup.title else ''
    name, city = extract_name_city(soup)
    experience = find_section_text(soup, SECTION_KEYWORDS['experience'])
    education = find_section_text(soup, SECTION_KEYWORDS['education'])
    skills_text = find_section_text(soup, SECTION_KEYWORDS['skills'])
    certifications = find_section_text(soup, SECTION_KEYWORDS['certifications'])
    skills_list, skill_years = extract_skills_and_years(skills_text)
    # also gather a short summary (first 3 paragraphs)
    paras = [p.get_text(separator=' ').strip() for p in soup.find_all('p') if p.get_text(strip=True)]
    summary = '\n'.join(paras[:3])
    # Detect technologies from the whole page text (French + English keywords)
    full_text = soup.get_text(separator=' ')
    technologies = detect_technologies(full_text)
    # join lists into semicolon-separated strings for single-cell Excel output
    skills_str = '; '.join(skills_list) if skills_list else ''
    tech_str = '; '.join(technologies) if technologies else ''
    tech_cat_str = categorize_technologies(technologies)
    return {
        'url': url,
        'page_title': title,
        'name': name,
        'city': city,
        'summary': summary,
        'experience': experience,
        'education': education,
        'skills_text': skills_text,
        'skills': skills_str,
        'technologies': tech_str,
        'technologies_categorized': tech_cat_str,
        'skill_years': json.dumps(skill_years, ensure_ascii=False),
        'certifications': certifications,
    }


def _normalize_sites(sites_cibles: Optional[str]) -> List[str]:
    if not sites_cibles:
        return []
    sites = []
    for raw in re.split(r"[,;\n]+", sites_cibles):
        site = raw.strip().strip('/').lower()
        if not site:
            continue
        if site.startswith('http://') or site.startswith('https://'):
            site = site.split('://', 1)[1]
        if site.startswith('www.'):
            site = site[4:]
        if site.startswith('site:'):
            site = site[5:]
        if not site:
            continue
        if site == 'linkedin.com':
            site = 'linkedin.com/in'
        if site == 'ictjob.com':
            site = 'ictjob.com'
        if '.' not in site:
            continue
        sites.append(site)
    return list(dict.fromkeys(sites))


def _build_base_terms(company: Optional[str], role: Optional[str], skills: Optional[str], experience: Optional[str], ville: Optional[str], pays: Optional[str] = None) -> str:
    parts = []
    if company:
        parts.append(f'"{company}"')
    if role:
        parts.append(f'"{role}"')
    if skills:
        for chunk in re.split(r"[,;/]+", skills):
            chunk = chunk.strip()
            if chunk:
                parts.append(f'"{chunk}"')
    if experience:
        for chunk in re.split(r"[,;/]+", experience):
            chunk = chunk.strip()
            if chunk:
                parts.append(f'"{chunk}"')
    if ville:
        parts.append(ville)
    if pays:
        parts.append(pays)
    if not parts:
        return 'IT data engineer'
    return ' '.join(parts)


def build_dorks(prenom: str, nom: str, ville: Optional[str], company: Optional[str] = None, role: Optional[str] = None, skills: Optional[str] = None, experience: Optional[str] = None, sites_cibles: Optional[str] = None, pays: Optional[str] = None) -> List[str]:
    """
    Construit des dorks ciblés sur les sites cibles (par défaut LinkedIn et ictjob.com).
    Les termes de recherche utilisent company/role/skills/experience/ville.
    """
    sites = _normalize_sites(sites_cibles) or ['linkedin.com/in', 'ictjob.com']
    base = _build_base_terms(company, role, skills, experience, ville, pays=pays)

    if prenom and nom:
        base_with_name = f'"{prenom} {nom}" {base}'.strip()
        return [f'site:{site} {base_with_name}' for site in sites]

    return [f'site:{site} {base}' for site in sites]


def _clean_text(value) -> str:
    if value is None:
        return ''
    if isinstance(value, float) and pd.isna(value):
        return ''
    return str(value).strip()


def process_rows(rows: List[Dict], api_key: str, output_xlsx: Optional[str] = None, max_results: int = 5, delay: float = 1.0) -> pd.DataFrame:
    if max_results <= 0:
        return pd.DataFrame()

    result_rows = []
    for person in rows:
        prenom = _clean_text(person.get('prenom') or person.get('first_name'))
        nom = _clean_text(person.get('nom') or person.get('last_name'))
        ville = _clean_text(person.get('ville') or person.get('city'))
        # optional filters from CSV: company (employer) and role
        company = _clean_text(person.get('company') or person.get('employer'))
        role = _clean_text(person.get('role') or person.get('poste'))
        skills = _clean_text(person.get('skills') or person.get('competences'))
        experience = _clean_text(person.get('experience') or person.get('experiences'))
        sites_cibles = _clean_text(person.get('sites_cibles') or person.get('sites'))
        pays = _clean_text(person.get('pays') or person.get('country'))
        # If no personal name is provided, proceed if company or role is present
        if not prenom and not nom and not (company or role or skills or experience or sites_cibles or pays):
            continue
        queries = build_dorks(
            prenom.strip(),
            nom.strip(),
            ville.strip(),
            company=company.strip() or None,
            role=role.strip() or None,
            skills=skills.strip() or None,
            experience=experience.strip() or None,
            sites_cibles=sites_cibles.strip() or None,
            pays=pays.strip() or None,
        )
        seen_urls = set()
        for q in queries:
            print(f"[search] {q}")
            try:
                results = search_serpapi(q, api_key, num=max_results)
            except Exception as e:
                print(f"Erreur SerpApi pour query '{q}': {e}")
                results = []
            for r in results:
                url = r.get('link')
                if not url or url in seen_urls:
                    continue
                seen_urls.add(url)
                html = fetch_page(url)
                if html:
                    parsed = parse_profile(html, url)
                else:
                    # fallback: build a minimal parsed record from SerpApi snippet / rich_snippet
                    title = r.get('title') or ''
                    snippet = r.get('snippet') or ''
                    rich = r.get('rich_snippet') or {}
                    city = ''
                    try:
                        top_ext = rich.get('top', {}).get('extensions', [])
                        if top_ext:
                            # choose first plausible extension (often contains location or role)
                            city = top_ext[0]
                    except Exception:
                        city = ''
                    parsed = {
                        'url': url,
                        'page_title': title,
                        'name': '',
                        'city': city,
                        'summary': snippet,
                        'experience': snippet,
                        'education': '',
                        'skills_text': '',
                        'skills': '',
                        'technologies': '',
                        'technologies_categorized': '',
                        'skill_years': json.dumps({}, ensure_ascii=False),
                        'certifications': ''
                    }
                parsed.update({
                    'query': q,
                    'prenom': prenom,
                    'nom': nom,
                    'ville_input': ville,
                    'company': company,
                    'role': role,
                    'skills_input': skills,
                    'experience_input': experience,
                    'sites_cibles': sites_cibles,
                    'pays': pays,
                    'status': 'new',
                })
                result_rows.append(parsed)
                # courte pause
                time.sleep(delay)
            # pause courte entre queries
            time.sleep(delay)

    if not result_rows:
        print('Aucun résultat trouvé.')
        return pd.DataFrame()

    df = pd.DataFrame(result_rows)
    cols = ['company', 'role', 'skills_input', 'experience_input', 'sites_cibles', 'ville_input', 'pays', 'query', 'url', 'page_title', 'city', 'experience', 'skills', 'technologies', 'technologies_categorized', 'skill_years', 'certifications', 'summary', 'status']
    available = [c for c in cols if c in df.columns]
    if available:
        df = df[available]
    if output_xlsx:
        df.to_excel(output_xlsx, index=False)
        print(f"Résultats sauvés dans {output_xlsx}")
    return df


def process_names(input_csv: str, api_key: str, output_xlsx: str, max_results: int = 5, delay: float = 1.0):
    with open(input_csv, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        rows = list(reader)
    process_rows(rows, api_key, output_xlsx=output_xlsx, max_results=max_results, delay=delay)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--api-key', help='Clé SerpApi. Si absente, lit SERPAPI_API_KEY env var')
    parser.add_argument('--input', default='input_names.csv', help='CSV d''entrée (prenom,nom,ville)')
    parser.add_argument('--output', default='results.xlsx', help='Fichier Excel de sortie')
    parser.add_argument('--max-results', type=int, default=5, help='Nombre de résultats organiques à récupérer par dork')
    parser.add_argument('--delay', type=float, default=1.0, help='Pause (s) entre requêtes pour réduire le rate')

    args = parser.parse_args()
    api_key = args.api_key or os.getenv('SERPAPI_API_KEY')
    if not api_key:
        print('Erreur: clé SerpApi manquante. Passez --api-key ou définissez la variable d''environnement SERPAPI_API_KEY')
        return
    process_names(args.input, api_key, args.output, max_results=args.max_results, delay=args.delay)


if __name__ == '__main__':
    main()
