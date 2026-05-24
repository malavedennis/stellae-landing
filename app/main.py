import io
import os
import re
import unicodedata
from datetime import datetime
from pathlib import Path

import anthropic
import hashlib
import pytesseract
from PIL import Image
from pdf2image import convert_from_bytes
import pandas as pd
import streamlit as st
from dotenv import load_dotenv
from docx import Document
from fpdf import FPDF
from PyPDF2 import PdfReader
from supabase import Client, create_client


def normalize_text(text):
    """
    Normaliza texto para matching bilingüe: minúsculas y sin acentos/diacríticos.
    Permite que keywords en español o inglés coincidan en contenido mixto.
    """
    nfkd = unicodedata.normalize("NFKD", str(text).lower())
    return "".join(c for c in nfkd if not unicodedata.combining(c))


# =============================================================================
# CONFIGURACIÓN INICIAL
# =============================================================================

PROJECT_ROOT = Path(__file__).resolve().parent.parent
load_dotenv(PROJECT_ROOT / ".env")

SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")

MAX_CHARS = 50_000

# Rutas de herramientas OCR — detección automática según sistema operativo
import sys as _sys
import shutil as _shutil

if _sys.platform == "win32":
    # Windows local — rutas hardcodeadas
    TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    POPPLER_PATH   = r"C:\Archivos de programa\poppler\Library\bin"
else:
    # Linux / Railway — binarios instalados via packages.txt (apt)
    TESSERACT_PATH = _shutil.which("tesseract") or "tesseract"
    POPPLER_PATH   = None  # pdf2image encuentra poppler automáticamente en Linux

# Configurar pytesseract
try:
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH
except Exception:
    pass
MODEL_ID = "claude-sonnet-4-6"

AUTHORITY_LEVEL_LABELS = {
    1: "1 — Field Engineer (solo lectura)",
    2: "2 — Project Engineer (puede reportar)",
    3: "3 — Project Manager (aprueba cambios menores)",
    4: "4 — Project Director (aprueba cambios mayores)",
    5: "5 — Steering Committee (aprobación máxima)",
}

CATEGORY_DISPLAY = {
    "decision": "🛡️ Decisiones Huérfanas",
    "change": "🔄 Cambios Ciegos",
    "risk": "⚠️ Riesgos Ocultos",
}

# Idiomas disponibles para el output del análisis
OUTPUT_LANGUAGES = {
    "English":    "en",
    "Español":    "es",
    "Português":  "pt",
    "Français":   "fr",
}

LANGUAGE_INSTRUCTIONS = {
    "en": "CRITICAL LANGUAGE RULE: You MUST write your ENTIRE response in ENGLISH, regardless of the language of the source documents. All findings, titles, descriptions, alerts, and labels must be in English.",
    "es": "REGLA DE IDIOMA CRÍTICA: Debes escribir tu respuesta COMPLETA en ESPAÑOL, independientemente del idioma de los documentos fuente. Todos los hallazgos, títulos, descripciones, alertas y etiquetas deben estar en español.",
    "pt": "REGRA DE IDIOMA CRÍTICA: Você DEVE escrever sua resposta COMPLETA em PORTUGUÊS, independentemente do idioma dos documentos fonte. Todos os achados, títulos, descrições, alertas e rótulos devem estar em português.",
    "fr": "RÈGLE DE LANGUE CRITIQUE: Vous DEVEZ écrire toute votre réponse en FRANÇAIS, quelle que soit la langue des documents source. Tous les résultats, titres, descriptions, alertes et étiquettes doivent être en français.",
}

# Etiquetas de estructura del system prompt por idioma
STRUCTURE_LABELS = {
    "en": {
        "role": "Act as a Principal Risk and Governance Auditor specialized in infrastructure, energy, and capital megaprojects (with the analytical rigor needed to prevent catastrophic failures like Berlin Brandenburg Airport or Crossrail). Your goal is to scan the documents provided by the user (minutes, reports, contracts, correspondence) and extract in a raw, objective, and corporate-jargon-free manner only three types of latent findings. Maintain a tone of professional skepticism: look for what parties are trying to omit, soften, or delegate informally.",
        "critical_rule": "CRITICAL STRUCTURE RULE: Each finding must be a complete and indivisible unit. The same change or decision MUST NOT generate multiple separate findings. If a change has several consequences, all must be grouped under a single finding with numbered sub-alerts. This is mandatory.",
        "orphan_label": "DECISION",
        "change_label": "CHANGE",
        "risk_label": "RISK",
        "finding": "Finding",
        "source_doc": "Source Document",
        "gov_impact": "Governance Impact",
        "mitigation": "Mitigation Action",
        "change_detected": "Change Detected",
        "cross_consequences": "Cross Consequences",
        "alert": "ALERT",
        "financial_exposure": "Financial Exposure Level",
        "alert_signal": "Alert Signal",
        "evidence": "Evidence",
        "escalation_consequences": "Consequences if Not Escalated",
        "days_30": "30 days",
        "days_60": "60 days",
        "inaction_cost": "Cost of Inaction",
        "no_findings": "No anomalies detected under current governance parameters in this section.",
        "steering": "Steering Committee",
        "hours_72": "next 48-72 hours",
    },
    "es": {
        "role": "Actúa como un Auditor Principal de Riesgos y Gobernanza especializado en megaproyectos de infraestructura, energía y capital (con el rigor analítico necesario para evitar fallas catastróficas como las del Aeropuerto de Berlín-Brandenburgo o Crossrail). Tu objetivo es escanear los documentos provistos por el usuario (minutas, reportes, contratos, correspondencia) y extraer de forma cruda, objetiva y sin adornos corporativos únicamente tres tipos de hallazgos latentes. Es crucial que asumas un tono de escepticismo profesional: busca lo que las partes intentan omitir, suavizar o delegar de manera informal.",
        "critical_rule": "REGLA DE ESTRUCTURA CRÍTICA: Cada hallazgo debe ser una unidad completa e indivisible. Un mismo cambio o decisión NO debe generar múltiples hallazgos separados. Si un cambio tiene varias consecuencias, todas deben estar agrupadas bajo un solo hallazgo con sub-alertas numeradas. Esto es mandatorio.",
        "orphan_label": "DECISIÓN",
        "change_label": "CAMBIO",
        "risk_label": "RIESGO",
        "finding": "Hallazgo",
        "source_doc": "Documento de Origen",
        "gov_impact": "Impacto en Gobernanza",
        "mitigation": "Acción de Mitigación",
        "change_detected": "Cambio Detectado",
        "cross_consequences": "Consecuencias Cruzadas",
        "alert": "ALERTA",
        "financial_exposure": "Nivel de Exposición Financiera",
        "alert_signal": "Señal de Alerta",
        "evidence": "Evidencia",
        "escalation_consequences": "Consecuencias si No Se Escala",
        "days_30": "A 30 días",
        "days_60": "A 60 días",
        "inaction_cost": "Costo de la Inacción",
        "no_findings": "No se detectaron anomalías bajo los parámetros de gobernanza actuales en esta sección.",
        "steering": "Comité Directivo",
        "hours_72": "próximas 48-72 horas",
    },
    "pt": {
        "role": "Aja como um Auditor Principal de Riscos e Governança especializado em megaprojetos de infraestrutura, energia e capital. Seu objetivo é escanear os documentos fornecidos pelo usuário e extrair de forma crua e objetiva apenas três tipos de achados latentes.",
        "critical_rule": "REGRA DE ESTRUTURA CRÍTICA: Cada achado deve ser uma unidade completa e indivisível. A mesma mudança ou decisão NÃO deve gerar múltiplos achados separados. Esto é mandatório.",
        "orphan_label": "DECISÃO",
        "change_label": "MUDANÇA",
        "risk_label": "RISCO",
        "finding": "Achado",
        "source_doc": "Documento de Origem",
        "gov_impact": "Impacto na Governança",
        "mitigation": "Ação de Mitigação",
        "change_detected": "Mudança Detectada",
        "cross_consequences": "Consequências Cruzadas",
        "alert": "ALERTA",
        "financial_exposure": "Nível de Exposição Financeira",
        "alert_signal": "Sinal de Alerta",
        "evidence": "Evidência",
        "escalation_consequences": "Consequências se Não Escalado",
        "days_30": "30 dias",
        "days_60": "60 dias",
        "inaction_cost": "Custo da Inação",
        "no_findings": "Nenhuma anomalia detectada sob os parâmetros de governança atuais nesta seção.",
        "steering": "Comitê Diretivo",
        "hours_72": "próximas 48-72 horas",
    },
    "fr": {
        "role": "Agissez en tant qu'Auditeur Principal des Risques et de la Gouvernance spécialisé dans les mégaprojets d'infrastructure, d'énergie et de capital. Votre objectif est de scanner les documents fournis et d'extraire uniquement trois types de constats latents.",
        "critical_rule": "RÈGLE DE STRUCTURE CRITIQUE: Chaque constat doit être une unité complète et indivisible. Le même changement ou décision NE DOIT PAS générer plusieurs constats séparés. C'est obligatoire.",
        "orphan_label": "DÉCISION",
        "change_label": "CHANGEMENT",
        "risk_label": "RISQUE",
        "finding": "Constat",
        "source_doc": "Document Source",
        "gov_impact": "Impact sur la Gouvernance",
        "mitigation": "Action de Mitigation",
        "change_detected": "Changement Détecté",
        "cross_consequences": "Conséquences Croisées",
        "alert": "ALERTE",
        "financial_exposure": "Niveau d'Exposition Financière",
        "alert_signal": "Signal d'Alerte",
        "evidence": "Preuves",
        "escalation_consequences": "Conséquences si Non Escaladé",
        "days_30": "30 jours",
        "days_60": "60 jours",
        "inaction_cost": "Coût de l'Inaction",
        "no_findings": "Aucune anomalie détectée sous les paramètres de gouvernance actuels dans cette section.",
        "steering": "Comité de Pilotage",
        "hours_72": "prochaines 48-72 heures",
    },
}

BASE_SYSTEM_PROMPT = """Actúa como un Auditor Principal de Riesgos y Gobernanza especializado en megaproyectos de infraestructura, energía y capital (con el rigor analítico necesario para evitar fallas catastróficas como las del Aeropuerto de Berlín-Brandenburgo o Crossrail). Tu objetivo es escanear los documentos provistos por el usuario (minutas, reportes, contratos, correspondencia) y extraer de forma cruda, objetiva y sin adornos corporativos únicamente tres tipos de hallazgos latentes. Es crucial que asumas un tono de escepticismo profesional: busca lo que las partes intentan omitir, suavizar o delegar de manera informal.

REGLA DE ESTRUCTURA CRÍTICA: Cada hallazgo debe ser una unidad completa e indivisible. Un mismo cambio o decisión NO debe generar múltiples hallazgos separados. Si un cambio tiene varias consecuencias, todas deben estar agrupadas bajo un solo hallazgo con sub-alertas numeradas. Esto es mandatorio.

Debes estructurar tu respuesta de manera estricta utilizando las siguientes etiquetas XML:

<decisiones_huerfanas>
Por cada decisión pendiente, aprobada informalmente o en el limbo, usa EXACTAMENTE este formato:

DECISIÓN [N]: [Título breve de la decisión]
- Hallazgo: Descripción precisa de qué se decidió o qué está en el limbo.
- Documento de Origen: Nombre del archivo, punto o tema específico donde aparece.
- Impacto en Gobernanza: Por qué la ausencia de responsable formal pone en riesgo el proyecto.
- Acción de Mitigación: Qué debe exigir el Steering Committee en las próximas 48-72 horas.

[Separar cada decisión con una línea en blanco]
</decisiones_huerfanas>

<cambios_ciegos>
Por cada cambio de alcance, diseño, ingeniería o contrato que carezca de análisis formal, usa EXACTAMENTE este formato:

CAMBIO [N]: [Título breve del cambio]
- Cambio Detectado: Qué se modificó, cómo y con qué justificación informal.
- Consecuencias Cruzadas:
  ALERTA 1 — [Área afectada, ej: Permisos Regulatorios]: Descripción del impacto específico.
  ALERTA 2 — [Área afectada, ej: Contratos con Terceros]: Descripción del impacto específico.
  ALERTA 3 — [Área afectada, ej: Sistemas Integrados]: Descripción del impacto específico.
  ALERTA 4 — [Área afectada, ej: Ruta Crítica]: Descripción del impacto específico.
  [Agregar todas las alertas relevantes — no limitar el número]
- Nivel de Exposición Financiera: ALTA / MEDIA / BAJA — con justificación cuantitativa estimada.

[Separar cada cambio con una línea en blanco]
</cambios_ciegos>

<riesgos_ocultos>
Por cada patrón de fricción o near-miss que el equipo esté normalizando, usa EXACTAMENTE este formato:

RIESGO [N]: [Título breve del riesgo]
- Señal de Alerta: El patrón detectado y por qué es una señal de alerta temprana.
- Evidencia: Referencias específicas a los documentos que demuestran recurrencia o patrón.
- Consecuencias si No Se Escala:
  1. A 30 días: Qué ocurre si no se actúa en el próximo mes.
  2. A 60 días: Qué ocurre si se sigue ignorando.
- Costo de la Inacción: Estimación del impacto económico o de cronograma si explota.

[Separar cada riesgo con una línea en blanco]
</riesgos_ocultos>

Regla estricta: Si no encuentras hallazgos en alguna categoría, escribe: 'No se detectaron anomalías bajo los parámetros de gobernanza actuales en esta sección'. No inventes datos. No fragmentes un mismo problema en múltiples hallazgos."""


def get_system_prompt(language_code: str = "en") -> str:
    """Construye el system prompt completo en el idioma seleccionado."""
    lang = LANGUAGE_INSTRUCTIONS.get(language_code, LANGUAGE_INSTRUCTIONS["en"])
    L = STRUCTURE_LABELS.get(language_code, STRUCTURE_LABELS["en"])

    prompt = f"""{lang}

{L["role"]}

{L["critical_rule"]}

You must structure your response strictly using the following XML tags:

<decisiones_huerfanas>
For each pending, informally approved, or limbo decision, use EXACTLY this format:

{L["orphan_label"]} [N]: [Brief title]
- {L["finding"]}: Precise description of what was decided or what is in limbo.
- {L["source_doc"]}: File name, specific point or topic where it appears.
- {L["gov_impact"]}: Why the absence of a formal owner puts the project at risk.
- {L["mitigation"]}: What the {L["steering"]} must demand in the {L["hours_72"]}.

[Separate each finding with a blank line]
</decisiones_huerfanas>

<cambios_ciegos>
For each scope, design, engineering, or contract change lacking formal analysis, use EXACTLY this format:

{L["change_label"]} [N]: [Brief title]
- {L["change_detected"]}: What was modified, how, and with what informal justification.
- {L["cross_consequences"]}:
  {L["alert"]} 1 — [Affected area]: Specific impact description.
  {L["alert"]} 2 — [Affected area]: Specific impact description.
  {L["alert"]} 3 — [Affected area]: Specific impact description.
  [Add all relevant alerts — do not limit the number]
- {L["financial_exposure"]}: HIGH / MEDIUM / LOW — with estimated quantitative justification.

[Separate each change with a blank line]
</cambios_ciegos>

<riesgos_ocultos>
For each friction pattern or near-miss the team is normalizing, use EXACTLY this format:

{L["risk_label"]} [N]: [Brief title]
- {L["alert_signal"]}: The detected pattern and why it is an early warning signal.
- {L["evidence"]}: Specific references to documents showing recurrence or pattern.
- {L["escalation_consequences"]}:
  1. {L["days_30"]}: What happens if no action is taken in the next month.
  2. {L["days_60"]}: What happens if it continues to be ignored.
- {L["inaction_cost"]}: Estimated economic or schedule impact if it explodes.

[Separate each risk with a blank line]
</riesgos_ocultos>

Strict rule: If you find no findings in a category, write: '{L["no_findings"]}'. Do not invent data. Do not fragment the same problem into multiple findings."""

    return prompt


# ======


# =============================================================================
# FUNCIONES AUXILIARES — EXTRACCIÓN DE TEXTO
# =============================================================================

def extract_text_with_ocr(image) -> str:
    """Extrae texto de una imagen usando OCR con pytesseract."""
    try:
        text = pytesseract.image_to_string(
            image,
            lang="spa+eng",  # Español e inglés
            config="--psm 3"  # Detección automática de layout
        )
        return text.strip()
    except Exception as e:
        return ""


def extract_text_from_pdf(uploaded_file) -> str:
    """
    Extrae texto de un PDF. Si el PDF es escaneado (sin texto digital),
    aplica OCR automáticamente página por página.
    """
    pdf_bytes = uploaded_file.getvalue()
    reader = PdfReader(io.BytesIO(pdf_bytes))

    pages_text = []
    ocr_needed_pages = []

    # Primera pasada: extraer texto digital
    for i, page in enumerate(reader.pages):
        page_text = page.extract_text()
        if page_text and len(page_text.strip()) > 30:
            pages_text.append((i, page_text.strip()))
        else:
            ocr_needed_pages.append(i)

    # Si hay páginas sin texto digital, aplicar OCR
    if ocr_needed_pages:
        try:
            _pdf2img_kwargs = dict(
                dpi=200,
                first_page=min(ocr_needed_pages) + 1,
                last_page=max(ocr_needed_pages) + 1,
            )
            if POPPLER_PATH:
                _pdf2img_kwargs["poppler_path"] = POPPLER_PATH
            images = convert_from_bytes(pdf_bytes, **_pdf2img_kwargs)
            for idx, page_num in enumerate(ocr_needed_pages):
                if idx < len(images):
                    ocr_text = extract_text_with_ocr(images[idx])
                    if ocr_text:
                        pages_text.append((page_num, f"[OCR] {ocr_text}"))
        except Exception as ocr_error:
            # Si OCR falla, continuar con el texto digital disponible
            pages_text.append((999, f"[OCR unavailable for some pages: {ocr_error}]"))

    # Ordenar por número de página y unir
    pages_text.sort(key=lambda x: x[0])
    return "\n".join(text for _, text in pages_text)


def extract_text_from_docx(uploaded_file) -> str:
    doc = Document(io.BytesIO(uploaded_file.getvalue()))
    paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
    return "\n".join(paragraphs)


def extract_text_from_txt(uploaded_file) -> str:
    return uploaded_file.getvalue().decode("utf-8", errors="replace")


def extract_text_from_image(uploaded_file) -> str:
    """Extrae texto de una imagen (JPG, PNG) usando OCR."""
    try:
        image = Image.open(io.BytesIO(uploaded_file.getvalue()))
        text = extract_text_with_ocr(image)
        return text if text else "[No readable text detected in image]"
    except Exception as e:
        return f"[Image OCR error: {e}]"


def extract_all_text(uploaded_files) -> str:
    """
    Extrae texto de todos los archivos subidos.
    Soporta: PDF (digital + escaneado con OCR), DOCX, TXT, JPG, PNG.
    """
    sections = []
    for uploaded_file in uploaded_files:
        extension = uploaded_file.name.rsplit(".", 1)[-1].lower()
        if extension == "pdf":
            text = extract_text_from_pdf(uploaded_file)
        elif extension == "docx":
            text = extract_text_from_docx(uploaded_file)
        elif extension == "txt":
            text = extract_text_from_txt(uploaded_file)
        elif extension in ("jpg", "jpeg", "png", "tiff", "bmp"):
            text = extract_text_from_image(uploaded_file)
        else:
            continue
        if text.strip():
            sections.append(f"--- Documento: {uploaded_file.name} ---\n{text}")
    return "\n\n".join(sections)


# =============================================================================
# FUNCIONES AUXILIARES — PROCESAMIENTO DE RESPUESTA XML
# =============================================================================

def normalize_alerts_formatting(text: str) -> str:
    """Post-procesamiento: garantiza que cada ALERT/ALERTA/ALERTE empiece en nueva línea.
    Elimina la variabilidad de formato de Claude — funciona en todos los idiomas."""
    if not text:
        return text

    # Patrones de etiquetas de alerta en los 4 idiomas soportados
    alert_patterns = [
        r'(ALERT\s+\d+\s+—)',      # English: ALERT 1 —
        r'(ALERTA\s+\d+\s+—)',     # Spanish: ALERTA 1 —
        r'(ALERTE\s+\d+\s+—)',     # French:  ALERTE 1 —
        r'(ALERTA\s+\d+\s+—)',     # Portuguese: same as Spanish
        r'(ALERT\s+\d+\s+-)',       # Variante con guion simple
        r'(ALERTA\s+\d+\s+-)',
    ]

    # También normalizar etiquetas de decisión, cambio y riesgo
    section_patterns = [
        r'(DECISION\s+\[?\d+\]?\s*:)',
        r'(CHANGE\s+\[?\d+\]?\s*:)',
        r'(RISK\s+\[?\d+\]?\s*:)',
        r'(DECISIÓN\s+\[?\d+\]?\s*:)',
        r'(CAMBIO\s+\[?\d+\]?\s*:)',
        r'(RIESGO\s+\[?\d+\]?\s*:)',
        r'(DÉCISION\s+\[?\d+\]?\s*:)',
        r'(CHANGEMENT\s+\[?\d+\]?\s*:)',
        r'(RISQUE\s+\[?\d+\]?\s*:)',
        r'(DECISÃO\s+\[?\d+\]?\s*:)',
        r'(MUDANÇA\s+\[?\d+\]?\s*:)',
        r'(RISCO\s+\[?\d+\]?\s*:)',
    ]

    result = text
    # Asegurar doble salto de línea antes de cada alerta
    newline = "\n"
    for pattern in alert_patterns:
        result = re.sub(
            r"(?<!" + newline + r")\s*" + pattern,
            newline + newline + r"",
            result,
            flags=re.IGNORECASE
        )
    # Asegurar doble salto de línea antes de cada hallazgo principal
    for pattern in section_patterns:
        result = re.sub(
            r"(?<!" + newline + r")\s*" + pattern,
            newline + newline + r"",
            result,
            flags=re.IGNORECASE
        )
    # Limpiar más de 3 saltos de línea consecutivos
    result = re.sub(r'\n{4,}', '\n\n\n', result)
    return result.strip()


def extract_tag(text, tag):
    pattern = f"<{tag}>(.*?)</{tag}>"
    match = re.search(pattern, text, re.DOTALL)
    if not match:
        return "No se detectaron anomalías bajo los parámetros de gobernanza actuales en esta sección."
    # Aplicar normalización de formato antes de devolver
    return normalize_alerts_formatting(match.group(1).strip())


def calculate_documents_hash(uploaded_files) -> str:
    """Calcula un hash MD5 único basado en el contenido de todos los archivos subidos."""
    hasher = hashlib.md5()
    for uploaded_file in sorted(uploaded_files, key=lambda f: f.name):
        hasher.update(uploaded_file.name.encode("utf-8"))
        hasher.update(uploaded_file.getvalue())
    return hasher.hexdigest()


def check_duplicate_analysis(supabase_client, project_id: str, doc_hash: str):
    """
    Verifica si el conjunto de documentos ya fue analizado para este proyecto.
    Retorna el análisis anterior si existe, None si es nuevo.
    """
    response = (
        supabase_client.table("analyses")
        .select("id, created_at, documents_analyzed")
        .eq("project_id", project_id)
        .eq("document_hash", doc_hash)
        .order("created_at", desc=True)
        .limit(1)
        .execute()
    )
    if response.data:
        return response.data[0]
    return None


def get_file_signature(uploaded_files):
    if not uploaded_files:
        return None
    return tuple((f.name, len(f.getvalue())) for f in uploaded_files)


def split_individual_findings(text: str) -> list:
    """
    Divide el bloque de texto en hallazgos individuales COMPLETOS.
    Reconoce el formato nuevo: CAMBIO [N]:, DECISION [N]:, RIESGO [N]:
    así como formatos legacy. Cada hallazgo incluye TODAS sus sub-secciones.
    """
    if not text or len(text.strip()) < 20:
        return []
    no_anomalies = "No se detectaron anomalías bajo los parámetros de gobernanza actuales"
    if no_anomalies in text:
        return []

    # Formato nuevo — separar por CAMBIO N:, DECISION N:, RIESGO N:, HALLAZGO N:
    nuevo_patron = r"(?=(?:CAMBIO|DECISI[OÓ]N|RIESGO|HALLAZGO)\s+\d+\s*:)"
    parts = re.split(nuevo_patron, text, flags=re.IGNORECASE | re.MULTILINE)
    fragments = [p.strip() for p in parts if p.strip() and len(p.strip()) >= 50]
    if len(fragments) > 1:
        return fragments

    # Formato legacy con numeración: "Hallazgo 1:", "Cambio Detectado 1:"
    legacy_patterns = [
        r"(?=Hallazgo\s+\d+\s*:)",
        r"(?=Cambio\s+Detectado\s+\d+\s*:)",
        r"(?=Se\u00f1al\s+de\s+Alerta\s+\d+\s*:)",
    ]
    for pattern in legacy_patterns:
        parts = re.split(pattern, text, flags=re.IGNORECASE | re.MULTILINE)
        fragments = [p.strip() for p in parts if p.strip() and len(p.strip()) >= 50]
        if len(fragments) > 1:
            return fragments

    # Si no hay patrones reconocibles, el bloque entero es un solo hallazgo
    full_text = text.strip()
    if len(full_text) >= 20:
        return [full_text]

    return []


def require_supabase_data(response, context: str) -> list:
    if not response.data:
        raise Exception(f"Empty response from Supabase ({context})")
    return response.data


# =============================================================================
# FUNCIONES AUXILIARES — GOVERNANCE CORE
# =============================================================================

def load_projects_map(supabase_client: Client) -> dict:
    response = supabase_client.table("projects").select("id, name").execute()
    if not response.data:
        return {}
    return {p["name"]: p["id"] for p in response.data}


def parse_roles_from_text(text: str) -> list:
    """
    Parsea roles desde texto extraído de un documento Word/Excel/PDF.
    Busca patrones como: Role Name | Level | Can Approve | Max Impact
    o tablas con esos encabezados. Retorna lista de dicts listos para insertar.
    """
    roles = []
    lines = [l.strip() for l in text.split("\n") if l.strip()]

    for line in lines:
        # Separadores comunes: | ; , tab
        for sep in ["|", ";", "\t"]:
            if sep in line:
                parts = [p.strip() for p in line.split(sep)]
                if len(parts) >= 2:
                    # Ignorar líneas de encabezado
                    header_words = ["role", "name", "level", "authority", "approval", "impact", "max"]
                    if any(w in parts[0].lower() for w in header_words):
                        continue
                    try:
                        role_name = parts[0]
                        if not role_name or len(role_name) < 2:
                            continue
                        # Extraer nivel de autoridad (buscar número 1-5)
                        authority_level = 3  # default
                        for p in parts[1:]:
                            nums = re.findall(r"\b([1-5])\b", p)
                            if nums:
                                authority_level = int(nums[0])
                                break
                        # Detectar si puede aprobar
                        can_approve = False
                        for p in parts:
                            if any(w in p.lower() for w in ["yes", "si", "true", "x", "approve"]):
                                can_approve = True
                                break
                        # Extraer monto máximo
                        max_impact = 0.0
                        for p in parts:
                            amounts = re.findall(r"[\d,\.]+", p.replace(",", ""))
                            for a in amounts:
                                try:
                                    val = float(a)
                                    if val > 100:  # ignora números pequeños como niveles
                                        max_impact = val
                                        break
                                except ValueError:
                                    pass
                        roles.append({
                            "role_name": role_name,
                            "authority_level": authority_level,
                            "can_approve_changes": can_approve,
                            "max_impact_value": max_impact,
                        })
                    except Exception:
                        continue
                break
    return roles


def parse_rules_from_text(text: str) -> list:
    """
    Parsea reglas de escalación desde texto extraído de un documento.
    Busca patrones como: Rule Name | Category | Keywords | Required Level
    Retorna lista de dicts listos para insertar.
    """
    rules = []
    lines = [l.strip() for l in text.split("\n") if l.strip()]

    for line in lines:
        for sep in ["|", ";", "\t"]:
            if sep in line:
                parts = [p.strip() for p in line.split(sep)]
                if len(parts) >= 3:
                    header_words = ["rule", "name", "category", "keyword", "level", "required"]
                    if any(w in parts[0].lower() for w in header_words):
                        continue
                    try:
                        rule_name = parts[0]
                        if not rule_name or len(rule_name) < 2:
                            continue
                        # Detectar categoría
                        category = "change"  # default
                        for p in parts:
                            p_lower = p.lower()
                            if "decision" in p_lower or "decisi" in p_lower:
                                category = "decision"
                                break
                            elif "risk" in p_lower or "riesgo" in p_lower:
                                category = "risk"
                                break
                            elif "change" in p_lower or "cambio" in p_lower:
                                category = "change"
                                break
                        # Extraer keywords (tercer campo o buscar comas)
                        keywords = []
                        if len(parts) >= 3:
                            kw_raw = parts[2]
                            keywords = [k.strip() for k in kw_raw.replace(";", ",").split(",") if k.strip()]
                        # Extraer nivel requerido
                        required_level = 5  # default SC
                        for p in parts:
                            nums = re.findall(r"\b([1-5])\b", p)
                            if nums:
                                required_level = int(nums[0])
                        if keywords:
                            rules.append({
                                "rule_name": rule_name,
                                "category": category,
                                "trigger_keywords": keywords,
                                "required_authority_level": required_level,
                                "is_active": True,
                            })
                    except Exception:
                        continue
                break
    return rules


def auto_generate_governance(document_text: str, project_name: str) -> dict:
    """
    Usa Claude para extraer roles y reglas de escalación desde un documento
    de proyecto (PEP, contrato, Project Charter, RACI, etc.).
    Retorna dict con 'roles' y 'rules' listos para confirmar e insertar.
    """
    client = anthropic.Anthropic()

    prompt = f"""Eres un experto en gobernanza de proyectos de capital en Oil & Gas, 
minería e infraestructura. Analiza el siguiente documento de proyecto y extrae:

1. La jerarquía de roles y niveles de autoridad que aparecen explícita o implícitamente
2. Las áreas o temas que según el documento requieren aprobación formal o escalación

Responde ÚNICAMENTE con un JSON válido con esta estructura exacta, sin texto adicional:
{{
  "roles": [
    {{
      "role_name": "nombre del rol",
      "authority_level": número del 1 al 5,
      "can_approve_changes": true o false,
      "max_impact_value": número en USD o 0 si no se especifica
    }}
  ],
  "rules": [
    {{
      "rule_name": "descripción corta de la regla",
      "category": "decision" o "change" o "risk",
      "trigger_keywords": ["keyword1", "keyword2", "keyword3"],
      "required_authority_level": número del 1 al 5
    }}
  ]
}}

Criterios para asignar authority_level:
1 = roles de campo / técnicos sin autoridad de aprobación
2 = ingenieros / especialistas que reportan hallazgos
3 = project managers / supervisores que aprueban cambios menores
4 = project directors / gerentes que aprueban cambios mayores
5 = steering committee / board / comité ejecutivo

Criterios para trigger_keywords:
- Usa palabras que aparecerían en documentos de proyecto cuando ese tema esté presente
- Incluye versiones en español e inglés de cada keyword relevante
- Mínimo 3 keywords por regla, máximo 8
- Palabras simples y específicas, no frases largas

Proyecto: {project_name}

Documento a analizar:
{document_text[:12000]}

Responde SOLO con el JSON. Sin explicaciones. Sin markdown. Sin texto antes o después."""

    message = client.messages.create(
        model=MODEL_ID,
        max_tokens=4000,
        messages=[{"role": "user", "content": prompt}]
    )

    response_text = message.content[0].text.strip()

    # Limpiar posibles backticks de markdown
    if response_text.startswith("```"):
        response_text = re.sub(r"```(?:json)?", "", response_text).strip()

    import json
    result = json.loads(response_text)
    return result


def load_project_context(supabase_client: Client, project_id: str) -> dict:
    """Carga el contexto completo de un proyecto desde Supabase."""
    response = (
        supabase_client.table("projects")
        .select("*")
        .eq("id", project_id)
        .execute()
    )
    if response.data:
        return response.data[0]
    return {}


def build_findings_list(decisiones_text: str, cambios_text: str, riesgos_text: str) -> list:
    category_texts = {
        "decision": decisiones_text,
        "change": cambios_text,
        "risk": riesgos_text,
    }
    findings_list = []
    for category, text in category_texts.items():
        for content in split_individual_findings(text):
            findings_list.append({
                "content": content,
                "category": category,
                "governance_violation": False,
                "violated_rule": None,
                "status": "open",
            })
    return findings_list


def flatten_findings(findings_by_category: dict) -> list:
    all_findings = []
    for findings in findings_by_category.values():
        all_findings.extend(findings)
    return all_findings


def rebuild_findings_by_category(all_findings: list) -> dict:
    findings_by_category = {"decision": [], "change": [], "risk": []}
    for finding in all_findings:
        findings_by_category[finding["category"]].append(finding)
    return findings_by_category


def apply_governance_rules(project_id: str, findings_list: list, supabase_client: Client) -> list:
    """Cruza hallazgos con escalation_rules activas usando matching bilingüe normalizado."""
    try:
        roles_response = (
            supabase_client.table("governance_roles")
            .select("id")
            .eq("project_id", project_id)
            .execute()
        )
        rules_response = (
            supabase_client.table("escalation_rules")
            .select("*")
            .eq("project_id", project_id)
            .eq("is_active", True)
            .execute()
        )
    except Exception as load_error:
        st.warning(f"⚠️ Could not load governance rules — analysis continues without violation detection: {load_error}")
        return findings_list

    if not roles_response.data:
        st.info("ℹ️ No governance rules defined for this project. Set them up in the Governance page to enable violation detection.")
        return findings_list

    active_rules = rules_response.data or []
    if not active_rules:
        st.info("ℹ️ No active escalation rules for this project. Add rules in the Governance page to enable violation detection.")
        return findings_list

    for finding in findings_list:
        finding["governance_violation"] = False
        finding["violated_rule"] = None
        content_norm = normalize_text(finding["content"])

        for rule in active_rules:
            if rule.get("category") != finding["category"]:
                continue
            keywords = rule.get("trigger_keywords") or []
            if isinstance(keywords, str):
                keywords = [k.strip() for k in keywords.split(",") if k.strip()]
            for keyword in keywords:
                keyword_norm = normalize_text(keyword)
                if keyword_norm and keyword_norm in content_norm:
                    finding["governance_violation"] = True
                    finding["violated_rule"] = rule.get("rule_name", "Unknown rule")
                    break
            if finding["governance_violation"]:
                break

    return findings_list


def render_category_findings_tab(findings_list: list, category: str) -> None:
    """Renderiza hallazgos con código de colores por categoría y violación."""
    if not findings_list:
        st.caption("No individual findings detected in this category.")
        st.caption("✅ No governance violations detected in this category")
        return

    for finding in findings_list:
        content = finding["content"]
        if finding.get("governance_violation"):
            violated_rule = finding.get("violated_rule") or "Unknown rule"
            st.error("🚨 GOVERNANCE VIOLATION — Escalation required: " + violated_rule)
            st.markdown(content)
            st.divider()
        elif category == "decision":
            st.markdown("⚠️ " + content)
            st.divider()
        elif category == "change":
            st.markdown("🔄 " + content)
            st.divider()
        elif category == "risk":
            st.markdown("🔴 " + content)
            st.divider()

    total = len(findings_list)
    violations = sum(1 for f in findings_list if f.get("governance_violation"))
    if violations > 0:
        st.caption(f"🚨 {violations} of {total} findings require mandatory escalation")
    else:
        st.caption("✅ No governance violations detected in this category")


def get_all_active_findings_flat() -> list:
    findings_by_category = st.session_state.get("findings_by_category")
    if not findings_by_category:
        return []
    return flatten_findings(findings_by_category)


def render_governance_sidebar_summary() -> None:
    """Mini-resumen en sidebar con conteos de gobernanza."""
    st.divider()
    st.markdown("**Governance summary**")

    if st.session_state.get("decisiones") is None:
        st.markdown("ℹ️ Run an analysis to see governance status")
        return

    all_findings = get_all_active_findings_flat()
    if not all_findings:
        st.markdown("ℹ️ Run an analysis to see governance status")
        return

    violations = sum(1 for f in all_findings if f.get("governance_violation"))
    in_review = sum(1 for f in all_findings if f.get("status") == "in_review")
    closed = sum(1 for f in all_findings if f.get("status") == "closed")

    if violations == 0 and in_review == 0 and closed == 0:
        st.markdown("ℹ️ Run an analysis to see governance status")
        return

    if violations > 0:
        st.markdown(f"🔴 {violations} governance violations")
    if in_review > 0:
        st.markdown(f"🟡 {in_review} findings in review")
    if closed > 0:
        st.markdown(f"🟢 {closed} findings closed")


# =============================================================================
# FUNCIONES AUXILIARES — SUPABASE (PERSISTENCIA)
# =============================================================================

def get_or_create_project(supabase_client: Client, project_name: str, 
                           project_description: str, industry: str = None,
                           project_type: str = None, project_stage: str = None,
                           context_document: str = None, context_filename: str = None) -> str:
    """Busca proyecto por nombre o lo crea con contexto completo."""
    response = supabase_client.table("projects").select("id").eq("name", project_name).execute()
    if response.data and len(response.data) > 0:
        project_id = response.data[0]["id"]
        # Actualizar contexto si se proporcionó documento nuevo
        if context_document:
            supabase_client.table("projects").update({
                "industry": industry,
                "project_type": project_type,
                "project_stage": project_stage,
                "context_document": context_document,
                "context_filename": context_filename,
            }).eq("id", project_id).execute()
        return project_id
    insert_response = supabase_client.table("projects").insert({
        "name": project_name,
        "description": project_description,
        "industry": industry,
        "project_type": project_type,
        "project_stage": project_stage,
        "context_document": context_document,
        "context_filename": context_filename,
    }).execute()
    if not insert_response.data or len(insert_response.data) == 0:
        raise Exception("Failed to create project in Supabase")
    return insert_response.data[0]["id"]


def save_findings_list(supabase_client: Client, analysis_id: str, findings_list: list) -> None:
    for finding in findings_list:
        supabase_client.table("findings").insert({
            "analysis_id": analysis_id,
            "category": finding["category"],
            "content": finding["content"],
            "status": finding.get("status", "open"),
            "governance_violation": finding.get("governance_violation", False),
            "violated_rule": finding.get("violated_rule"),
        }).execute()


def save_analysis_to_supabase(supabase_client: Client, project_id: str, document_names: list,
                               char_count: int, raw_output: str, findings_by_category: dict,
                               doc_hash: str = None) -> None:
    # Calcular hash del conjunto de documentos para detección de duplicados futuros
    response = supabase_client.table("analyses").insert({
        "project_id": project_id,
        "documents_analyzed": document_names,
        "characters_processed": char_count,
        "raw_output": raw_output,
        "document_hash": doc_hash if doc_hash else None,
    }).execute()
    data = require_supabase_data(response, "insert analysis")
    analysis_id = data[0]["id"]
    all_findings = flatten_findings(findings_by_category)
    save_findings_list(supabase_client, analysis_id, all_findings)

    # Pre-cachear traducción al inglés en background si el análisis no está en inglés
    # Esto hace que el primer Export PDF en inglés sea instantáneo
    try:
        current_lang = st.session_state.get("output_language", "en")
        if current_lang != "en":
            # Recargar findings con IDs para poder guardar el cache
            fresh = supabase_client.table("findings").select("*").eq("analysis_id", analysis_id).execute()
            if fresh.data:
                translate_findings_for_pdf(fresh.data, "en")
    except Exception:
        pass  # No bloquear si falla el pre-cache


def format_analysis_date(created_at: str) -> str:
    if not created_at:
        return "Fecha desconocida"
    try:
        normalized = created_at.replace("Z", "+00:00")
        dt = datetime.fromisoformat(normalized)
        return dt.strftime("%d/%m/%Y %H:%M")
    except (ValueError, TypeError):
        return str(created_at)[:16]


def clean_markdown_for_table(text: str) -> str:
    """Elimina símbolos Markdown para mostrar texto limpio en tablas."""
    if not text:
        return ""
    # Eliminar negritas y cursivas
    text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
    text = re.sub(r'\*(.+?)\*', r'\1', text)
    # Eliminar encabezados markdown
    text = re.sub(r'^#{1,6}\s+', '', text, flags=re.MULTILINE)
    # Eliminar guiones de listas al inicio de línea
    text = re.sub(r'^\s*-\s+', '', text, flags=re.MULTILINE)
    # Limpiar espacios múltiples
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()


def render_findings_editor(supabase_client: Client, analysis_id: str) -> None:
    response = supabase_client.table("findings").select("*").eq("analysis_id", analysis_id).execute()
    if not response.data:
        st.caption("No findings recorded for this analysis.")
        return

    df_original = pd.DataFrame(response.data)

    # Crear copia limpia para visualización — sin asteriscos Markdown
    df_display = df_original.copy()
    if "content" in df_display.columns:
        df_display["content"] = df_display["content"].apply(clean_markdown_for_table)

    edited_df = st.data_editor(
        df_display,
        column_order=["category", "content", "status"],
        column_config={
            "category": st.column_config.TextColumn("Category", disabled=True),
            "content": st.column_config.TextColumn("Content", disabled=True, width="large"),
            "status": st.column_config.SelectboxColumn("Status", options=["open", "in_review", "closed"]),
        },
        hide_index=True,
        key=f"findings_editor_{analysis_id}",
    )

    for _, row in edited_df.iterrows():
        finding_id = row["id"]
        original_row = df_original[df_original["id"] == finding_id]
        if original_row.empty:
            continue
        if row["status"] != original_row.iloc[0]["status"]:
            supabase_client.table("findings").update({"status": row["status"]}).eq("id", finding_id).execute()
            st.toast("✅ Status updated")


# =============================================================================
# FUNCIONES AUXILIARES — DASHBOARD EJECUTIVO (SEMANA 6)
# =============================================================================

def load_all_project_findings(supabase_client: Client, project_id: str) -> list:
    """Carga todos los findings de todos los análisis de un proyecto."""
    analyses_response = (
        supabase_client.table("analyses")
        .select("id")
        .eq("project_id", project_id)
        .execute()
    )
    if not analyses_response.data:
        return []

    all_findings = []
    for analysis in analyses_response.data:
        findings_response = (
            supabase_client.table("findings")
            .select("*")
            .eq("analysis_id", analysis["id"])
            .execute()
        )
        if findings_response.data:
            all_findings.extend(findings_response.data)
    return all_findings


def calculate_project_status(all_findings: list) -> tuple:
    """
    Calcula el semáforo del proyecto.
    Retorna (status_label, status_message, status_level)
    status_level: 'critical', 'at_risk', 'ok'
    """
    total_violations = sum(1 for f in all_findings if f.get("governance_violation"))
    total_open = sum(1 for f in all_findings if f.get("status") == "open")

    if total_violations > 0:
        return (
            "🔴 CRITICAL",
            f"{total_violations} governance violation(s) require immediate escalation",
            "critical"
        )
    elif total_open > 5:
        return (
            "🟡 AT RISK",
            f"{total_open} open findings require attention",
            "at_risk"
        )
    else:
        return (
            "🟢 UNDER CONTROL",
            "No critical governance issues detected",
            "ok"
        )


def clean_for_pdf(text: str) -> str:
    """Limpia texto para compatibilidad con fpdf2 — elimina caracteres Unicode no soportados."""
    if not text:
        return ""
    replacements = {
        "\u2014": "-",    # guión largo —
        "\u2013": "-",    # guión medio –
        "\u2018": "'",    # comilla izquierda '
        "\u2019": "'",    # comilla derecha '
        "\u201c": '"',    # comilla doble izquierda "
        "\u201d": '"',    # comilla doble derecha "
        "\u2022": "-",    # bullet •
        "\u2026": "...",  # ellipsis …
        "\u00b0": " grados",
        "\u00e9": "e", "\u00f3": "o", "\u00fa": "u",
        "\u00ed": "i", "\u00e1": "a", "\u00f1": "n",
        "\u00c9": "E", "\u00d3": "O", "\u00da": "U",
        "\u00cd": "I", "\u00c1": "A", "\u00d1": "N",
        "\u00fc": "u", "\u00e0": "a", "\u00e8": "e", "\u00f2": "o",
        "\u00bf": "?", "\u00a1": "!",
        "\u20ac": "EUR", "\u00a3": "GBP",
    }
    for char, replacement in replacements.items():
        text = text.replace(char, replacement)
    # Eliminar emojis y cualquier carácter fuera de ASCII
    return text.encode("ascii", errors="replace").decode("ascii").replace("?", " ").strip()


def strip_markdown(text: str) -> str:
    """Elimina sintaxis Markdown del texto para el PDF."""
    if not text:
        return ""
    # Eliminar negritas **texto**
    import re as _re
    text = _re.sub(r"\*\*(.+?)\*\*", r"\1", text)
    # Eliminar cursivas *texto*
    text = _re.sub(r"\*(.+?)\*", r"\1", text)
    # Eliminar encabezados markdown
    text = _re.sub(r"^#{1,6}\s+", "", text, flags=_re.MULTILINE)
    # Normalizar listas
    text = _re.sub(r"^\s*[-*]\s+", "- ", text, flags=_re.MULTILINE)
    # Normalizar saltos de línea múltiples
    text = _re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()

def pdf_section_header(pdf: FPDF, title: str, color: tuple = (30, 60, 114)) -> None:
    """Encabezado de sección con fondo de color."""
    pdf.set_fill_color(*color)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(0, 10, f"  {title}", ln=True, fill=True)
    pdf.set_text_color(0, 0, 0)
    pdf.ln(3)


def pdf_kpi_box(pdf: FPDF, label: str, value: str, color: tuple) -> None:
    """Caja de KPI con color de fondo."""
    pdf.set_fill_color(*color)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 18)
    pdf.cell(44, 16, value, align="C", fill=True)
    pdf.set_font("Helvetica", "", 8)
    pdf.set_text_color(80, 80, 80)
    x = pdf.get_x() - 44
    pdf.set_xy(x, pdf.get_y() + 16)
    pdf.cell(44, 6, label, align="C")
    pdf.set_xy(pdf.get_x(), pdf.get_y() - 22)


# Etiquetas del PDF por idioma
PDF_LABELS = {
    "en": {
        "title": "STELLAE - Executive Governance Report",
        "project": "Project",
        "generated": "Generated",
        "confidential": "CONFIDENTIAL",
        "status": "Project Status",
        "findings_summary": "Findings Summary",
        "orphan_decisions": "Orphan Decisions",
        "blind_changes": "Blind Changes",
        "hidden_risks": "Hidden Risks",
        "open": "Open",
        "in_review": "In Review",
        "closed": "Closed",
        "recent_analyses": "Recent Analyses",
        "documents": "Documents",
        "findings": "Findings",
        "governance_violation": "GOVERNANCE VIOLATION",
        "footer": "Generated by Stellae Governance Intelligence  |  Confidential — For Steering Committee Use Only",
    },
    "es": {
        "title": "STELLAE - Reporte Ejecutivo de Gobernanza",
        "project": "Proyecto",
        "generated": "Generado",
        "confidential": "CONFIDENCIAL",
        "status": "Estado del Proyecto",
        "findings_summary": "Resumen de Hallazgos",
        "orphan_decisions": "Decisiones Huérfanas",
        "blind_changes": "Cambios Ciegos",
        "hidden_risks": "Riesgos Ocultos",
        "open": "Abierto",
        "in_review": "En Revisión",
        "closed": "Cerrado",
        "recent_analyses": "Análisis Recientes",
        "documents": "Documentos",
        "findings": "Hallazgos",
        "governance_violation": "VIOLACIÓN DE GOBERNANZA",
        "footer": "Generado por Stellae Governance Intelligence  |  Confidencial — Solo para Comité Directivo",
    },
    "pt": {
        "title": "STELLAE - Relatório Executivo de Governança",
        "project": "Projeto",
        "generated": "Gerado",
        "confidential": "CONFIDENCIAL",
        "status": "Status do Projeto",
        "findings_summary": "Resumo de Achados",
        "orphan_decisions": "Decisões Órfãs",
        "blind_changes": "Mudanças Cegas",
        "hidden_risks": "Riscos Ocultos",
        "open": "Aberto",
        "in_review": "Em Revisão",
        "closed": "Fechado",
        "recent_analyses": "Análises Recentes",
        "documents": "Documentos",
        "findings": "Achados",
        "governance_violation": "VIOLAÇÃO DE GOVERNANÇA",
        "footer": "Gerado por Stellae Governance Intelligence  |  Confidencial — Apenas para Comitê Diretivo",
    },
    "fr": {
        "title": "STELLAE - Rapport Exécutif de Gouvernance",
        "project": "Projet",
        "generated": "Généré",
        "confidential": "CONFIDENTIEL",
        "status": "Statut du Projet",
        "findings_summary": "Résumé des Constats",
        "orphan_decisions": "Décisions Orphelines",
        "blind_changes": "Changements Aveugles",
        "hidden_risks": "Risques Cachés",
        "open": "Ouvert",
        "in_review": "En Révision",
        "closed": "Clôturé",
        "recent_analyses": "Analyses Récentes",
        "documents": "Documents",
        "findings": "Constats",
        "governance_violation": "VIOLATION DE GOUVERNANCE",
        "footer": "Généré par Stellae Governance Intelligence  |  Confidentiel — Réservé au Comité de Pilotage",
    },
}


def detect_content_language(text: str) -> str:
    """Detecta el idioma del texto — heurística rápida sin API."""
    sample = text[:500].lower()

    spanish_markers = [
        "hallazgo", "cambio detectado", "riesgo", "decisión", "alerta",
        "evidencia", "consecuencias", "gobernanza", "proyecto", " el ", " la ",
        " de ", " que ", " con ", " por ", " los ", " las ", "señal"
    ]
    english_markers = [
        "finding", "change detected", "risk", "decision", "alert",
        "evidence", "consequences", "governance", "project", " the ",
        " of ", " and ", " for ", " with ", " this ", "signal"
    ]
    portuguese_markers = [
        "achado", "mudança detectada", "risco", "decisão", "alerta",
        "evidência", "consequências", "governança", "projeto"
    ]
    french_markers = [
        "constat", "changement détecté", "risque", "décision", "alerte",
        "preuves", "conséquences", "gouvernance", "projet"
    ]

    scores = {
        "es": sum(1 for m in spanish_markers if m in sample),
        "en": sum(1 for m in english_markers if m in sample),
        "pt": sum(1 for m in portuguese_markers if m in sample),
        "fr": sum(1 for m in french_markers if m in sample),
    }
    best = max(scores, key=scores.get)
    return best if scores[best] >= 2 else "en"  # default inglés


def translate_single_batch(texts: list, target_lang: str) -> list:
    """Traduce una lista de textos con una sola llamada a Haiku."""
    lang_name = next((k for k, v in OUTPUT_LANGUAGES.items() if v == target_lang), "English")
    separator = "\n|||SEP|||\n"
    batch = separator.join(texts)
    try:
        client = anthropic.Anthropic()
        msg = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=8192,
            messages=[{
                "role": "user",
                "content": (
                    f"Translate the following governance findings to {lang_name}. "
                    "Each finding is separated by |||SEP|||. "
                    "Preserve ALL formatting, numbering, labels and structure exactly. "
                    "Do NOT translate proper nouns, file names, project names, or technical codes. "
                    "Return ONLY the translated text with the same separators, nothing else.\n\n"
                    f"{batch}"
                )
            }],
        )
        parts = msg.content[0].text.strip().split("|||SEP|||")
        return [p.strip() for p in parts]
    except Exception:
        return texts


def translate_findings_for_pdf(findings: list, target_lang: str) -> list:
    """Traduce findings al idioma del PDF. Usa cache en Supabase para evitar re-traducir."""
    if not findings:
        return findings

    # Detectar si el contenido ya está en el idioma target — skip si coincide
    sample = findings[0].get("content", "")
    detected = detect_content_language(sample)
    # DEBUG TEMPORAL — eliminar después
    import streamlit as _st
    _st.caption(f"🔍 DEBUG: target={target_lang} | detected={detected} | sample='{sample[:80]}'")
    if detected == target_lang:
        return findings  # Ya está en el idioma correcto — no traducir

    # Separar findings con y sin traducción cacheada
    needs_translation = []
    needs_translation_idx = []
    result = list(findings)  # copia para modificar

    for i, finding in enumerate(findings):
        cached = (finding.get("content_translations") or {}).get(target_lang)
        if cached:
            result[i] = dict(finding)
            result[i]["content"] = cached
        else:
            needs_translation.append(finding.get("content", ""))
            needs_translation_idx.append(i)

    # DEBUG TEMPORAL
    import streamlit as _st2
    _st2.caption(f"🔍 DEBUG2: needs_translation={len(needs_translation)} | cached={len(findings)-len(needs_translation)}")

    # Traducir solo los que no tienen cache
    if needs_translation:
        translated = translate_single_batch(needs_translation, target_lang)

        # Guardar en Supabase y actualizar result
        for j, idx in enumerate(needs_translation_idx):
            translated_text = translated[j] if j < len(translated) else needs_translation[j]
            result[idx] = dict(findings[idx])
            result[idx]["content"] = translated_text

            # Guardar en Supabase si el finding tiene ID
            finding_id = findings[idx].get("id")
            if finding_id:
                try:
                    existing = findings[idx].get("content_translations") or {}
                    existing[target_lang] = translated_text
                    supabase.table("findings").update(
                        {"content_translations": existing}
                    ).eq("id", finding_id).execute()
                except Exception:
                    pass  # No bloquear el PDF si falla el cache

    return result


# Logo PNG embebido en base64 para el PDF ejecutivo
_STELLAE_LOGO_PNG_B64 = "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAIAAABt+uBvAAAABmJLR0QA/wD/AP+gvaeTAAAKr0lEQVR4nO2ba3BU1R3A/+e+9r2bhLxfkIRHFIFIIKA8FBRSSweqjspYLKW1dhyd6Yxtpx38UNtObafvqa3TTrUtjtMnUEUqggbUhALNAwIoIRvyIsk+8tpkn/dx7umHlZhsNnuzu/fCdub+Pu2ePefcc3/7v/9zz7m7qPzuvaAzN9StHkCmowtSQBekgC5IAV2QArogBXRBCuiCFNAFKaALUkAXpIAuSAFdkAK6IAV0QQroghTIOEEIoVs9hBlknKAsh+1WD2EGGSeouCjvVg9hBpknqFAXlBBdkALFRXkZlaczTpDJaMioPJ1xgiDD8nRGCsqkNKSmIAqp01uagihVU5iagmiWtTuy0+8nnTxtd2TRrCH9MUyhpiBR4BFjrKysSLOflPN0aWkJoYyiEElzANNROQcFJsZ8YVSzajlN0+n0k2yepilq5fJqP8+EAuPpHHc2KgvCEs+Hgv3e8Mb1tRazMeV+kkpDFrNp3dqaXm9EFCNY5FM+aFzUn8UiQR/G8qWukU0b63IXZKXWScm8Iyg/L3v9utUfdY8hBHzQl9rhEqC+ICzxEh/GhJy9eP2uutUVC0tS6KSocF55eklV2coVy1suDwJCIh9SPXxAo/ugSGAcCBAZPmzruXPV7WvvXJ5sD/PJ05vvqqmqqmy+PAgIAdEkfEAjQRgLAh8AACBw8sy1svLiz27bmOzMnSBPI4Qe3LHZbM86c6E/WiLyQSwJaQx5TrS6kxZCPgIEAAgiJ886adaw+6HtLMvMv4e58jTLMk/ve3A8IJ1p64uWECCRkMqT1xRaCcKSKEYC0deEoMbWa5NB4am9Dzls1nn2EDdPZ9ltzz+39yOn6/yVIbgRkWIkKEuiGqOOg4ZrMT7oA0KirwlBTW29zl7XU/senucUPjtPl5cWfPfb+469337hqhum9BDCB7UKH9BUkIwl4UYQRTl3caCl3fnlPbtuW6p8tx2Tp1cuX/z8c3v/8sbpjp7R6dWEiF/Gklpjno22q3k+6AMg00varriamq/sfrh+y6Y1is2n8nT91nXPPb375QMnOvvGZtQghA9NqDfeOCSRNVNAliUh7OdM9umF568MiaK0/Z41WXbbkXc+xBjP1by4MK/jas8XHq3funH1j18+7BoJx1TQOnwAgHaU1Wh6ACyJnMkek028Y8GA379h3R0Ly4o6OnulaY62bFo79Vomcv2WtWtrql986bB7bNYSlJDQxDAhspbD137DjMiSGJ6cXX6py3vsVOui8uIHP3fvXG2L8rKXLS7/4UuHPeNxFuhCeFKWtQ0fuDk7inxwAuJ9z1e6R46caF66pKJmxbLZn46NjS2pKnnhF//0jkcAgKVlGn3aCSFyROPsE0XbHBRFJhikCWDj7KV1XR9768R/H9i24fqAe3T80xMOBSNLq8pe+PlBX1CkKHnP3f6ddZR/dPC6V+oasTlHrM4B7CdzJi8VQRr9mcVskMsXyNVFYnWxuLgQszT6ztEVISH+97GwyL56eflrfz0qE/L9/U8DEJ6P/PlgYyAs77wz8PhmbDYZAEAIB32enqlWYQH1jdAdLvbqENPpZifDmjwsUi2CaIoszMXVxVJVvlSZL5Vk45i11/1LvUcuF8dt2+ealOW+9WtW/qe5HQA8w6OHj7fVLvQ/uVXMdhinBsmZLJzROnVvZeJIdbFUXSxBLQDAeIjq9tDdXqbDxV51Mbyojq+0BBXYcXWJVJkvVeZJVQWYpUmCytuWeRo684NzBNF1j78035yflzM86jvfcvpHj4Qqik2AYrclLdn5gisQt4dss1xbIddWiABhLIPLR3d7mWtepmOI6R1m5ERDS0TSlxiNSH6WXGAjBXbMMskddiCY2zORaJ2RbZb3bA7cv4pLsPI/d9HrHYnvaC7CAtUzyjlddI+bSEmqSjqCMEGucdo1nkpbBCFbLkZUnO1qQmQs8Ysrql5544PaioIcBzdXJwvyC188xhKifAURIsuigCUBSzwWBYxT2Q/R/EZxJgQAGM40s4xIEk8hqN94e93q2442tF3r6llTnW02xf8Ccqyky00NjcW5QSFACJYkISxGAnzQFw6MCRG/JISxJJBUpzwNBJFPV9qzwZLAGa2Iip4ewSIPgGiaeXzn+tLiPJZlTzW2uPwWyXOuvLzIZokfR4vy4PiFTxIekSUsREQ+IIQmI/5RPjQR3XtV6x5SA0Hx7FAMR7MGkGVCZEDAcmYsCSATmjUgBF95dCNNgd1mA4BTjS0AqM+fzY025ReV2a1xHGVZSNdgxNnvC/tH+aBP5IOSGJGxGLMwVoWbcokhsGQVGswOg8XOGq0IEBCgWQ5RNCLyk7s3szRyOBzRuqcaWwBAxIyEUcTdmltY5rDFeVJanif/q9FHUp6c5s3NWGqwBivNRAMB0TTLGi00ywEAIviZJ+4rLcgxxHtY/PFoqS9saDj2zrXrcZYUiwqY+2pMs8tVR3NBCJDREufpGAL56/u2LyzLc3YPMhw7uwIBONa7kqOEpvdOOPviPLHY9xkrTWn+UyvNBTEmC0XHnj8C/I0nH8jJsr74ywOD7pG52vKYO+iswxjONhz/2Dkc82lJLlO/RvMg0lgQQkZz7BqVAvytr+1AIH/vp3/s7XcPuWLPfDqjYevBzroJwdDa1HC5M7bml+qtDKNtEGkryGC0UfSM2xmakvc/u8vjHv7Bz/40Nj4JAEPuRIIAIChxh51r+ycXtJ9+t/WyZ/pHBdn0jjptg0hLQQhxZsf0ApqS9z+zq/1S529ePcQLnzyocXtHZKywK4gJ9W7fHReGyzuaG1ouuqZ/9MXtNgOrYRBpOM0bzA7WaJl6y9Jk/7O73jz6wdETp8m02ZkQcnt1pdVqjr6NTvNxGQzkTApmS7A1KDvKij9Rbzag8YB8pf//7bkYQtT08OEY8s2v7vjDgTcbz7XPrjyYMA1Nx+krONq96uqFs03N/VOFe+6zGjmtgkgrQZzJRt1YlBo59MwT9//q93/rcPbFrZw4T8fgDmYddq6+dL614XR3tCTHTn1+gznNAc+FJoIQogw3wsfMocd21P7k16+7PaNz1VfM0zFMCuZDnWvazne+19gVLXl8q9Vk0CSINBHEme3RPQ2rmd62YdlvXzkUCif63eB88nQMosy83bPyg+bBf5/sAACHhXpkk0WxVQqoLwhRlMHkAACbiV5Rlfva349hxUkKy96RscR1ZiMT9OHA0rfPBN5q6ACAx7ZYbCb1T0f9Hg0mB6Iom5FaYEXvNJydZ6v55+kYLo+W/KNJOvKe02qiHr1H/SBSWRCiaIPZbjNRlBxo/8g5/4ZJ5ekY+icWHDgFh453PXKvOcui8hmp3J3B7DCyEJzw9g94lGtPI9k8HcNI2PrqSeat9wce26JyEKn6VwSKNps438jQuM+fbNsU8nQMQdHwu+N0KBDItql6Uir2xTD0+Ihnag2RFBjLnuGk83RsJ4R6/X3CIjV30dQUJAgCIakPLs2rLAoB8E5mqqA0SSdPa0cmCVIjglQngwSln6e1IIMEqZKnVSeDBEFGXmUZJijz8nSGCdIjKDGe4Tn3jG4VmSVIkm7Gzw6TIrMEZSC6IAV0QQroghTQBSmgC1JAF6SALkgBXZACuiAFdEEK6IIU0AUpoAtSQBekwP8AfM9VKGK+Y4QAAAAASUVORK5CYII="


def generate_executive_pdf(project_name: str, status_label: str, status_message: str,
                            all_findings: list, analyses: list,
                            language_code: str = "en") -> bytes:
    """Genera el PDF del reporte ejecutivo con diseño visual mejorado."""
    _raw_labels = PDF_LABELS.get(language_code, PDF_LABELS["en"])
    L = {k: clean_for_pdf(v) if isinstance(v, str) else v for k, v in _raw_labels.items()}
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.add_page()
    pdf.set_margins(15, 15, 15)

    # ── ENCABEZADO CON LOGO ─────────────────────────────────────────────────
    pdf.set_fill_color(15, 20, 25)
    pdf.rect(0, 0, 210, 36, 'F')
    pdf.set_text_color(255, 255, 255)

    # Logo PNG embebido
    try:
        import base64 as _b64, tempfile, os
        _logo_bytes = _b64.b64decode(_STELLAE_LOGO_PNG_B64)
        _tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
        _tmp.write(_logo_bytes)
        _tmp.close()
        pdf.image(_tmp.name, x=8, y=5, w=26, h=26)
        os.unlink(_tmp.name)
    except Exception:
        pass  # Si falla el logo, continúa sin él

    # Título
    pdf.set_font("Helvetica", "B", 16)
    pdf.set_xy(38, 7)
    pdf.cell(140, 10, clean_for_pdf(L["title"]), ln=False)

    # Gold accent line
    pdf.set_draw_color(201, 168, 76)
    pdf.set_line_width(0.5)
    pdf.line(38, 19, 200, 19)

    # Subtítulo
    pdf.set_font("Helvetica", "", 8)
    pdf.set_text_color(180, 180, 180)
    pdf.set_xy(38, 21)
    pdf.cell(170, 6, f'{L["project"]}: {clean_for_pdf(project_name)}   |   {L["generated"]}: {datetime.now().strftime("%d/%m/%Y %H:%M")}   |   {L["confidential"]}', ln=True)
    pdf.set_text_color(0, 0, 0)
    pdf.ln(12)

    # ── SECCIÓN 1: PROJECT STATUS ────────────────────────────────────────────
    status_text = status_label.replace("🔴", "").replace("🟡", "").replace("🟢", "").strip()
    if "CRITICAL" in status_text:
        status_color = (180, 30, 30)
        status_bg = (255, 235, 235)
        status_icon = "[!!] CRITICAL"
    elif "AT RISK" in status_text:
        status_color = (180, 100, 0)
        status_bg = (255, 248, 220)
        status_icon = "[!] AT RISK"
    else:
        status_color = (30, 120, 30)
        status_bg = (235, 255, 235)
        status_icon = "[OK] UNDER CONTROL"

    pdf_section_header(pdf, "1. PROJECT STATUS")
    pdf.set_fill_color(*status_bg)
    pdf.set_draw_color(*status_color)
    pdf.set_line_width(0.8)
    pdf.rect(15, pdf.get_y(), 180, 18, 'DF')
    pdf.set_text_color(*status_color)
    pdf.set_font("Helvetica", "B", 14)
    pdf.set_xy(20, pdf.get_y() + 3)
    pdf.cell(80, 8, clean_for_pdf(status_icon), ln=False)
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(60, 60, 60)
    pdf.cell(95, 8, clean_for_pdf(status_message), ln=True)
    pdf.set_draw_color(0, 0, 0)
    pdf.set_line_width(0.2)
    pdf.set_text_color(0, 0, 0)
    pdf.ln(8)

    # ── SECCIÓN 2: KPIs ──────────────────────────────────────────────────────
    total_violations = sum(1 for f in all_findings if f.get("governance_violation"))
    total_open = sum(1 for f in all_findings if f.get("status") == "open")
    total_in_review = sum(1 for f in all_findings if f.get("status") == "in_review")
    total_closed = sum(1 for f in all_findings if f.get("status") == "closed")

    pdf_section_header(pdf, "2. KPI SUMMARY")
    kpi_y = pdf.get_y()
    kpis = [
        ("Governance Violations", str(total_violations), (180, 30, 30)),
        ("Open Findings", str(total_open), (180, 100, 0)),
        ("In Review", str(total_in_review), (30, 100, 180)),
        ("Closed", str(total_closed), (30, 130, 30)),
    ]
    for i, (label, value, color) in enumerate(kpis):
        pdf.set_xy(15 + i * 46, kpi_y)
        pdf_kpi_box(pdf, label, value, color)
    pdf.set_xy(15, kpi_y + 24)
    pdf.ln(6)

    # ── SECCIÓN 3: BREAKDOWN ─────────────────────────────────────────────────
    pdf_section_header(pdf, "3. BREAKDOWN BY CATEGORY")
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_fill_color(240, 240, 240)
    pdf.cell(60, 8, "Category", border=1, fill=True, align="C")
    pdf.cell(40, 8, "Total Findings", border=1, fill=True, align="C")
    pdf.cell(40, 8, "Open", border=1, fill=True, align="C")
    pdf.cell(40, 8, "Violations", border=1, fill=True, ln=True, align="C")

    for cat_key, cat_label in [("decision", "Decisions"), ("change", "Changes"), ("risk", "Risks")]:
        cat_findings = [f for f in all_findings if f.get("category") == cat_key]
        cat_open = sum(1 for f in cat_findings if f.get("status") == "open")
        cat_viol = sum(1 for f in cat_findings if f.get("governance_violation"))
        pdf.set_font("Helvetica", "", 9)
        pdf.set_fill_color(255, 255, 255)
        if cat_viol > 0:
            pdf.set_text_color(180, 30, 30)
        pdf.cell(60, 7, cat_label, border=1, fill=True)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(40, 7, str(len(cat_findings)), border=1, fill=True, align="C")
        pdf.cell(40, 7, str(cat_open), border=1, fill=True, align="C")
        viol_text = str(cat_viol) if cat_viol == 0 else f"[!!] {cat_viol}"
        if cat_viol > 0:
            pdf.set_text_color(180, 30, 30)
        pdf.cell(40, 7, viol_text, border=1, fill=True, ln=True, align="C")
        pdf.set_text_color(0, 0, 0)
    pdf.ln(8)

    # ── SECCIÓN 4: GOVERNANCE VIOLATIONS ────────────────────────────────────
    violations_list = [f for f in all_findings if f.get("governance_violation")]
    if violations_list:
        pdf_section_header(pdf, "4. GOVERNANCE VIOLATIONS - IMMEDIATE ACTION REQUIRED", color=(160, 20, 20))
        for i, f in enumerate(violations_list, 1):
            # Cabecera de cada violation
            pdf.set_fill_color(255, 235, 235)
            pdf.set_draw_color(180, 30, 30)
            pdf.set_line_width(0.5)
            rule_clean = clean_for_pdf(f.get("violated_rule", "N/A"))
            cat = f["category"].upper()
            pdf.set_font("Helvetica", "B", 10)
            pdf.set_text_color(180, 30, 30)
            header_text = clean_for_pdf(f"  [{i}] [{cat}] -- Rule violated: {rule_clean}")
            pdf.cell(0, 8, header_text, ln=True, fill=True, border="LRB")
            pdf.set_text_color(0, 0, 0)
            pdf.set_line_width(0.2)
            pdf.set_draw_color(0, 0, 0)

            # Contenido completo sin truncar -- limpiado de Markdown
            raw_content = strip_markdown(clean_for_pdf(f["content"]))
            pdf.set_font("Helvetica", "", 9)
            pdf.set_fill_color(255, 248, 248)
            pdf.multi_cell(0, 5.5, raw_content, fill=True)
            pdf.ln(5)
    else:
        pdf_section_header(pdf, "4. GOVERNANCE VIOLATIONS", color=(30, 120, 30))
        pdf.set_font("Helvetica", "", 11)
        pdf.cell(0, 8, "No governance violations detected in this report.", ln=True)
        pdf.ln(5)

    # ── SECCIÓN 5: TOP OPEN FINDINGS ────────────────────────────────────────
    open_findings = [f for f in all_findings if f.get("status") == "open"]
    open_findings_sorted = sorted(open_findings, key=lambda x: x.get("created_at", ""))[:5]
    if open_findings_sorted:
        pdf_section_header(pdf, "5. OLDEST OPEN FINDINGS (TOP 5 -- PRIORITY ATTENTION)")
        cat_colors = {
            "decision": (220, 235, 255),
            "change": (220, 245, 220),
            "risk": (255, 235, 220),
        }
        for i, f in enumerate(open_findings_sorted, 1):
            date_str = format_analysis_date(f.get("created_at", ""))
            cat = f["category"]
            bg = cat_colors.get(cat, (240, 240, 240))
            pdf.set_fill_color(*bg)
            pdf.set_font("Helvetica", "B", 10)
            pdf.cell(0, 7, f"  [{i}] [{cat.upper()}] -- Detected: {date_str}", ln=True, fill=True)
            raw_content = strip_markdown(clean_for_pdf(f["content"]))
            pdf.set_font("Helvetica", "", 9)
            pdf.set_fill_color(250, 250, 250)
            pdf.multi_cell(0, 5.5, raw_content, fill=True)
            pdf.ln(4)

    # ── FOOTER ───────────────────────────────────────────────────────────────
    pdf.set_y(-12)
    pdf.set_fill_color(15, 20, 25)
    pdf.rect(0, pdf.get_y(), 210, 12, 'F')
    pdf.set_text_color(180, 180, 180)
    pdf.set_font("Helvetica", "I", 7)
    pdf.cell(0, 12, clean_for_pdf(L["footer"]), align="C")

    return bytes(pdf.output())


# =============================================================================
# FUNCIONES AUXILIARES -- ANÁLISIS CON CLAUDE
# =============================================================================

def init_session_state() -> None:
    defaults = {
        "decisiones": None,
        "cambios": None,
        "riesgos": None,
        "findings_by_category": None,
        "findings_list": None,
        "last_analyzed_signature": None,
        "analysis_meta": None,
        "project_name": "",
        "project_description": "",
        "project_id": None,
        "output_language": "en",   # Idioma de output por defecto: inglés
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def run_anthropic_analysis(document_text: str, project_context: dict = None) -> str:
    """Envía documentos a Claude con contexto del proyecto para análisis más preciso."""
    client = anthropic.Anthropic()

    # Construir el bloque de contexto del proyecto si existe
    context_block = ""
    if project_context:
        context_parts = []
        if project_context.get("name"):
            context_parts.append(f"Nombre del proyecto: {project_context['name']}")
        if project_context.get("description"):
            context_parts.append(f"Descripción: {project_context['description']}")
        if project_context.get("industry"):
            context_parts.append(f"Industria/Sector: {project_context['industry']}")
        if project_context.get("project_type"):
            context_parts.append(f"Tipo de proyecto: {project_context['project_type']}")
        if project_context.get("project_stage"):
            context_parts.append(f"Etapa actual: {project_context['project_stage']}")
        if project_context.get("context_document"):
            context_parts.append(
                f"\n--- DOCUMENTO DE REFERENCIA DEL PROYECTO ({project_context.get('context_filename', 'contexto')}) ---\n"
                f"{project_context['context_document'][:8000]}"
                f"\n--- FIN DEL DOCUMENTO DE REFERENCIA ---"
            )
        if context_parts:
            context_block = (
                "=== CONTEXTO DEL PROYECTO ===\n"
                + "\n".join(context_parts)
                + "\n=== FIN DEL CONTEXTO ===\n\n"
                "Usa este contexto para calibrar tu análisis. Detecta desviaciones del plan "
                "original, incumplimientos de los procedimientos establecidos, y decisiones "
                "que contradicen lo definido en el documento de referencia.\n\n"
            )

    user_message = (
        f"{context_block}"
        "Analiza los siguientes documentos de proyecto y "
        "responde estrictamente con las tres etiquetas XML solicitadas:\n\n"
        f"{document_text}"
    )

    message = client.messages.create(
        model=MODEL_ID,
        max_tokens=8192,
        system=get_system_prompt(st.session_state.get("output_language", "en")),
        messages=[{"role": "user", "content": user_message}],
    )
    return message.content[0].text


def parse_and_store_results(raw_response: str, num_docs: int, char_count: int) -> None:
    st.session_state.decisiones = extract_tag(raw_response, "decisiones_huerfanas")
    st.session_state.cambios = extract_tag(raw_response, "cambios_ciegos")
    st.session_state.riesgos = extract_tag(raw_response, "riesgos_ocultos")
    st.session_state.analysis_meta = {
        "num_docs": num_docs,
        "char_count": char_count,
        "model": MODEL_ID,
    }


def run_governance_pipeline(supabase_client: Client, project_id: str) -> dict:
    findings_list = build_findings_list(
        st.session_state.decisiones,
        st.session_state.cambios,
        st.session_state.riesgos,
    )
    findings_list = apply_governance_rules(project_id, findings_list, supabase_client)
    findings_by_category = rebuild_findings_by_category(findings_list)
    st.session_state.findings_by_category = findings_by_category
    st.session_state.findings_list = findings_list
    return findings_by_category


def translate_analysis_results(target_language_code: str) -> None:
    """Retraduce los resultados ya en session_state al idioma seleccionado — 3 llamadas paralelas."""
    decisiones = st.session_state.get("decisiones", "")
    cambios    = st.session_state.get("cambios", "")
    riesgos    = st.session_state.get("riesgos", "")

    if not decisiones and not cambios and not riesgos:
        return

    lang_name = next((k for k, v in OUTPUT_LANGUAGES.items() if v == target_language_code), "English")

    def translate_section(text: str) -> str:
        """Traduce una sección individual — prompt mínimo para máxima velocidad."""
        if not text or "No se detectaron" in text or "No findings" in text:
            return text
        client = anthropic.Anthropic()
        msg = client.messages.create(
            model="claude-haiku-4-5-20251001",  # Haiku: 5x más rápido para traducción simple
            max_tokens=4096,
            messages=[{
                "role": "user",
                "content": (
                    f"Translate the following governance analysis text COMPLETELY to {lang_name}. "
                    "Translate EVERYTHING including labels like 'Cambio Detectado', 'Hallazgo', "
                    "'Consecuencias Cruzadas', 'Señal de Alerta', 'Evidencia', 'Acción de Mitigación', "
                    "'Impacto en Gobernanza', 'Costo de la Inacción', 'Nivel de Exposición Financiera'. "
                    "Keep section headers like DECISIÓN [N], CAMBIO [N], RIESGO [N], ALERTA N in the target language. "
                    "Do NOT translate proper nouns, project names, file names, or company names. "
                    "Return ONLY the translated text, nothing else.\n\n"
                    f"{text}"
                )
            }],
        )
        return msg.content[0].text.strip()

    # Traducir las 3 secciones secuencialmente con Haiku (rápido)
    new_decisiones = translate_section(decisiones)
    new_cambios    = translate_section(cambios)
    new_riesgos    = translate_section(riesgos)

    # Actualizar session_state directamente
    st.session_state.decisiones = new_decisiones
    st.session_state.cambios    = new_cambios
    st.session_state.riesgos    = new_riesgos
    st.session_state.output_language = target_language_code

    # Reconstruir findings_by_category para que los tabs se actualicen
    findings_list = build_findings_list(new_decisiones, new_cambios, new_riesgos)
    project_id = st.session_state.get("project_id")
    if project_id:
        try:
            findings_list = apply_governance_rules(project_id, findings_list, supabase)
        except Exception:
            pass
    findings_by_category = rebuild_findings_by_category(findings_list)
    st.session_state.findings_by_category = findings_by_category
    st.session_state.findings_list = findings_list


def render_analysis_results_tabs() -> None:
    if st.session_state.get("decisiones") is None:
        return
    if st.session_state.get("findings_by_category") is None:
        return

    findings_by_category = st.session_state.findings_by_category

    # ── SELECTOR DE IDIOMA — traducción instantánea ─────────────
    lang_col, _ = st.columns([1, 3])
    with lang_col:
        lang_labels = list(OUTPUT_LANGUAGES.keys())
        current_code = st.session_state.get("output_language", "en")
        current_label = next((k for k, v in OUTPUT_LANGUAGES.items() if v == current_code), "English")
        selected_lang = st.selectbox(
            "🌐 Report language",
            options=lang_labels,
            index=lang_labels.index(current_label),
            key="lang_selector",
            help="Instantly translates the analysis results to the selected language.",
        )
        new_code = OUTPUT_LANGUAGES[selected_lang]
        if new_code != current_code:
            with st.spinner(f"Translating to {selected_lang}..."):
                translate_analysis_results(new_code)
            # Limpiar el widget para que refleje el nuevo idioma en el rerun
            if "lang_selector" in st.session_state:
                del st.session_state["lang_selector"]
            st.rerun()
    # ────────────────────────────────────────────────────────────

    tab_decisiones, tab_cambios, tab_riesgos = st.tabs([
        "🛡️ Decisiones Huérfanas",
        "🔄 Cambios Ciegos",
        "⚠️ Riesgos Ocultos",
    ])

    with tab_decisiones:
        render_category_findings_tab(findings_by_category.get("decision", []), "decision")
    with tab_cambios:
        render_category_findings_tab(findings_by_category.get("change", []), "change")
    with tab_riesgos:
        render_category_findings_tab(findings_by_category.get("risk", []), "risk")

    meta = st.session_state.analysis_meta
    if meta:
        st.caption(
            f"📊 Analysis complete — {meta['num_docs']} documents · {meta['char_count']:,} characters"
        )


# =============================================================================
# PÁGINAS DE LA APP (MULTI-PAGE)
# =============================================================================

def render_project_badge(project_name: str) -> None:
    """Muestra el badge del proyecto activo debajo del título de cada página."""
    if not project_name:
        return
    st.markdown(
        f'''<div style="
            display: inline-flex;
            align-items: center;
            gap: 6px;
            background: rgba(30,58,95,0.4);
            border: 1px solid rgba(201,168,76,0.3);
            border-radius: 20px;
            padding: 3px 12px 3px 8px;
            margin-bottom: 16px;
        ">
            <span style="color:#C9A84C; font-size:12px;">📁</span>
            <span style="color:#8899a6; font-size:12px; font-weight:500;">{project_name}</span>
        </div>''',
        unsafe_allow_html=True
    )


def _render_project_management_panel(project_id: str, project_name: str, supabase_client) -> None:
    """Panel de gestion del proyecto — Rename y Delete. Siempre visible en el Dashboard."""

    col_rename, col_spacer, col_danger = st.columns([2, 2, 1])

    # ── RENAME ──────────────────────────────────────────────────────────────
    with col_rename:
        if st.button("✏️ Rename project", use_container_width=True, key="btn_rename_project"):
            st.session_state.show_rename_form = not st.session_state.get("show_rename_form", False)

    if st.session_state.get("show_rename_form"):
        with st.form("rename_project_form"):
            new_name = st.text_input("New project name", value=project_name,
                                     placeholder="Enter new name")
            col_save, col_cancel = st.columns([1, 1])
            with col_save:
                save = st.form_submit_button("💾 Save", type="primary", use_container_width=True)
            with col_cancel:
                cancel = st.form_submit_button("Cancel", use_container_width=True)

        if save and new_name.strip() and new_name.strip() != project_name:
            try:
                supabase_client.table("projects").update(
                    {"name": new_name.strip()}
                ).eq("id", project_id).execute()
                st.session_state.project_name = new_name.strip()
                st.session_state.show_rename_form = False
                st.success(f"✅ Project renamed to '{new_name.strip()}'")
                st.rerun()
            except Exception as e:
                st.error(f"❌ Could not rename: {e}")
        elif cancel:
            st.session_state.show_rename_form = False
            st.rerun()

    # ── DELETE ───────────────────────────────────────────────────────────────
    with col_danger:
        if st.button("🗑️ Delete", type="secondary", use_container_width=True,
                     key="btn_delete_project"):
            st.session_state.delete_project_step1 = True
            st.session_state.delete_project_id = project_id
            st.session_state.delete_project_name = project_name
            st.rerun()

    # ── CONFIRMACIONES DE DELETE (dentro del panel, siempre accesibles) ───
    if st.session_state.get("delete_project_step1") and             st.session_state.get("delete_project_id") == project_id:
        proj_to_delete = st.session_state.get("delete_project_name", "")
        st.warning(
            f"⚠️ **First confirmation:** You are about to permanently delete "
            f"**{proj_to_delete}** and ALL its data — analyses, findings, roles "
            f"and escalation rules. This cannot be undone."
        )
        col_yes1, col_no1 = st.columns([1, 1])
        with col_yes1:
            if st.button("Yes, I want to delete this project",
                         key="del_proj_yes1", type="primary", use_container_width=True):
                st.session_state.delete_project_step2 = True
                st.rerun()
        with col_no1:
            if st.button("Cancel", key="del_proj_cancel1", use_container_width=True):
                for key in ["delete_project_step1", "delete_project_step2",
                            "delete_project_id", "delete_project_name"]:
                    if key in st.session_state:
                        del st.session_state[key]
                st.rerun()

    if st.session_state.get("delete_project_step2") and             st.session_state.get("delete_project_id") == project_id:
        proj_to_delete = st.session_state.get("delete_project_name", "")
        st.error("**Second confirmation:** Type the project name exactly to confirm deletion.")
        st.caption(f"Type exactly: **{proj_to_delete}**")
        confirm_name = st.text_input(
            "Type project name to confirm:",
            placeholder=proj_to_delete,
            key="delete_confirm_name_input"
        )
        delete_enabled = confirm_name.strip() == proj_to_delete.strip()
        col_del, col_cancel = st.columns([1, 1])
        with col_del:
            if st.button("🗑️ DELETE PERMANENTLY", key="del_proj_final",
                         type="primary", use_container_width=True,
                         disabled=not delete_enabled):
                try:
                    supabase_client.table("projects").delete().eq(
                        "id", project_id
                    ).execute()
                    for key in ["delete_project_step1", "delete_project_step2",
                                "delete_project_id", "delete_project_name",
                                "delete_confirm_name_input", "project_id",
                                "project_name", "project_description"]:
                        if key in st.session_state:
                            del st.session_state[key]
                    st.success("✅ Project deleted successfully.")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Failed to delete project: {e}")
        with col_cancel:
            if st.button("Cancel", key="del_proj_cancel2", use_container_width=True):
                for key in ["delete_project_step1", "delete_project_step2",
                            "delete_project_id", "delete_project_name"]:
                    if key in st.session_state:
                        del st.session_state[key]
                st.rerun()
        if not delete_enabled and confirm_name:
            st.caption("⚠️ Project name does not match. Check spelling and try again.")


def render_dashboard_page(supabase_client: Client) -> None:
    """Página Dashboard Ejecutivo -- semáforo, KPIs, timeline y export PDF."""
    st.title("📊 Executive Dashboard")
    render_project_badge(st.session_state.get("project_name", ""))
    st.markdown(
        '<p class="subtitle">Real-time governance status for your project</p>',
        unsafe_allow_html=True,
    )

    # Usar proyecto seleccionado globalmente en el sidebar
    project_name = st.session_state.get("project_name", "")
    project_id = st.session_state.get("project_id")
    selected_name = project_name

    if not project_id:
        st.info("ℹ️ Select a project in the sidebar to view the dashboard.")
        return

    # Cargar todos los análisis del proyecto
    analyses_response = (
        supabase_client.table("analyses")
        .select("*")
        .eq("project_id", project_id)
        .order("created_at", desc=True)
        .execute()
    )

    if not analyses_response.data:
        st.info("ℹ️ No analyses found for this project. Go to Analysis to run the first one.")

        # ── PANEL DE GESTION DEL PROYECTO (visible aunque no haya análisis) ──
        st.divider()
        _render_project_management_panel(project_id, selected_name, supabase_client)
        return

    analyses = analyses_response.data

    # Cargar todos los findings del proyecto
    all_findings = load_all_project_findings(supabase_client, project_id)

    # Calcular semáforo
    status_label, status_message, status_level = calculate_project_status(all_findings)

    # --- Sección 1: Semáforo prominente + botón de reporte al lado ---
    col_status, col_report = st.columns([3, 1])
    with col_status:
        st.markdown(f"## Project Status: {status_label}")
        st.markdown(f"*{status_message}*")
    with col_report:
        st.markdown("&nbsp;", unsafe_allow_html=True)
        # Selector de idioma para el PDF — guardado en session_state
        pdf_lang_options = list(OUTPUT_LANGUAGES.keys())
        default_pdf_lang = next(
            (k for k, v in OUTPUT_LANGUAGES.items()
             if v == st.session_state.get("output_language", "en")), "English"
        )
        pdf_lang_selected = st.selectbox(
            "🌐 PDF language",
            options=pdf_lang_options,
            index=pdf_lang_options.index(
                st.session_state.get("pdf_lang_selected", default_pdf_lang)
            ),
            key="pdf_lang_selector",
            label_visibility="collapsed",
        )
        st.session_state.pdf_lang_selected = pdf_lang_selected
        pdf_lang_code = OUTPUT_LANGUAGES[pdf_lang_selected]
        if st.button("📄 Export Report", type="primary", use_container_width=True):
            try:
                lang_label = next((k for k, v in OUTPUT_LANGUAGES.items() if v == pdf_lang_code), "English")
                spinner_msg = (
                    f"🌐 Translating findings to {lang_label} and generating PDF..."
                    if pdf_lang_code != "en" or detect_content_language(
                        all_findings[0].get("content", "") if all_findings else ""
                    ) == "es"
                    else "📄 Generating PDF report..."
                )
                with st.spinner(spinner_msg):
                    # Traducir findings al idioma del PDF si es necesario
                    pdf_findings = translate_findings_for_pdf(all_findings, pdf_lang_code)
                    pdf_bytes = generate_executive_pdf(
                        selected_name, status_label, status_message, pdf_findings, analyses,
                        language_code=pdf_lang_code,
                    )
                file_name = f"stellae_report_{selected_name.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
                st.download_button(
                    label="⬇️ Download PDF",
                    data=pdf_bytes,
                    file_name=file_name,
                    mime="application/pdf",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"❌ Failed: {e}")
    st.divider()

    # --- Sección 2: KPIs en tiempo real ---
    total_violations = sum(1 for f in all_findings if f.get("governance_violation"))
    total_open = sum(1 for f in all_findings if f.get("status") == "open")
    total_in_review = sum(1 for f in all_findings if f.get("status") == "in_review")
    total_closed = sum(1 for f in all_findings if f.get("status") == "closed")

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric(label="🚨 Governance Violations", value=total_violations)
    with col2:
        st.metric(label="⚠️ Open Findings", value=total_open)
    with col3:
        st.metric(label="🔄 In Review", value=total_in_review)
    with col4:
        st.metric(label="✅ Closed", value=total_closed)

    st.divider()

    # --- Sección 3: Breakdown por categoría ---
    st.markdown("### Breakdown by Category")
    col1, col2, col3 = st.columns(3)

    with col1:
        decisions = [f for f in all_findings if f.get("category") == "decision"]
        open_d = sum(1 for f in decisions if f.get("status") == "open")
        viol_d = sum(1 for f in decisions if f.get("governance_violation"))
        st.markdown("**🛡️ Decisions**")
        st.markdown(f"- Total: {len(decisions)}")
        st.markdown(f"- Open: {open_d}")
        st.markdown(f"- Violations: {viol_d}")

    with col2:
        changes = [f for f in all_findings if f.get("category") == "change"]
        open_c = sum(1 for f in changes if f.get("status") == "open")
        viol_c = sum(1 for f in changes if f.get("governance_violation"))
        st.markdown("**🔄 Changes**")
        st.markdown(f"- Total: {len(changes)}")
        st.markdown(f"- Open: {open_c}")
        st.markdown(f"- Violations: {viol_c}")

    with col3:
        risks = [f for f in all_findings if f.get("category") == "risk"]
        open_r = sum(1 for f in risks if f.get("status") == "open")
        viol_r = sum(1 for f in risks if f.get("governance_violation"))
        st.markdown("**⚠️ Risks**")
        st.markdown(f"- Total: {len(risks)}")
        st.markdown(f"- Open: {open_r}")
        st.markdown(f"- Violations: {viol_r}")

    st.divider()

    # --- Sección 4: Timeline de análisis -- máximo 10, con scroll interno ---
    st.markdown("### 📅 Analysis History")
    analyses_limited = analyses[:10]
    if len(analyses) > 10:
        st.caption(f"Showing last 10 of {len(analyses)} analyses.")

    with st.container(height=500):
        for analysis in analyses_limited:
            fecha_str = format_analysis_date(analysis.get("created_at", ""))
            docs = analysis.get("documents_analyzed", [])
            chars = analysis.get("characters_processed", 0)

            # Cargar todos los findings de este análisis incluyendo contenido completo
            findings_response = (
                supabase_client.table("findings")
                .select("*")
                .eq("analysis_id", analysis["id"])
                .execute()
            )
            findings_all = findings_response.data or []
            violations_in_analysis = [f for f in findings_all if f.get("governance_violation")]
            violations_count = len(violations_in_analysis)

            violation_label = f"🚨 {violations_count} violations" if violations_count > 0 else "✅ No violations"

            with st.expander(f"📋 {fecha_str} -- {len(docs)} doc(s) -- {violation_label}"):
                st.markdown(f"**Documents:** {', '.join(docs)}")
                st.markdown(f"**Characters processed:** {chars:,}")
                st.markdown(f"**Findings:** {len(findings_all)} total, {violations_count} violations")

                st.markdown("---")

                # Mostrar violations detalladas dentro del expander
                if violations_count > 0:
                    st.markdown(f"**🚨 Governance Violations ({violations_count})**")
                    for finding in violations_in_analysis:
                        st.error(
                            "🚨 " + finding.get("violated_rule", "Unknown rule") +
                            " -- [" + finding["category"].upper() + "]"
                        )
                        st.markdown(finding["content"])
                        st.divider()
                else:
                    st.caption("✅ No governance violations in this analysis.")

    # --- Zona de gestión — siempre visible al final del Dashboard ---
    st.divider()
    _render_project_management_panel(project_id, selected_name, supabase_client)




def render_analysis_page(supabase_client: Client) -> None:
    """Página principal: upload, análisis con Claude y tabs de resultados."""
    st.title("Stellae -- Governance Intelligence")
    render_project_badge(st.session_state.get("project_name", ""))
    st.markdown(
        '<p class="subtitle">Upload your project documents for analysis</p>',
        unsafe_allow_html=True,
    )

    uploaded_files = st.file_uploader(
        "Choose project documents",
        type=["pdf", "docx", "txt", "jpg", "jpeg", "png", "tiff", "bmp"],
        accept_multiple_files=True,
        help="Supported formats: PDF (digital or scanned), DOCX, TXT, JPG, PNG -- OCR enabled",
    )

    current_signature = get_file_signature(uploaded_files)

    if (
        current_signature is not None
        and st.session_state.last_analyzed_signature is not None
        and current_signature != st.session_state.last_analyzed_signature
    ):
        st.session_state.decisiones = None
        st.session_state.cambios = None
        st.session_state.riesgos = None
        st.session_state.findings_by_category = None
        st.session_state.findings_list = None
        st.session_state.analysis_meta = None

    if uploaded_files:
        st.subheader("Uploaded files")
        for uploaded_file in uploaded_files:
            size_kb = len(uploaded_file.getvalue()) / 1024
            st.markdown(
                f"""
                <div class="file-card">
                    <span class="file-name">{uploaded_file.name}</span>
                    <span class="file-size">{size_kb:.1f} KB</span>
                </div>
                """,
                unsafe_allow_html=True,
            )

    # ── Aviso de duplicado pendiente (persiste entre reruns) ──────────────────
    if st.session_state.get("pending_duplicate_info"):
        st.warning(st.session_state.pending_duplicate_info)
        col_rerun, col_cancel = st.columns([1, 1])
        with col_rerun:
            if st.button("🔄 Yes, analyze again", type="primary",
                         key="btn_force_reanalysis", use_container_width=True):
                st.session_state.force_reanalysis = True
                st.session_state.auto_run_analysis = True  # arranca el análisis automáticamente
                del st.session_state["pending_duplicate_info"]
                st.rerun()
        with col_cancel:
            if st.button("❌ Cancel", key="btn_cancel_reanalysis", use_container_width=True):
                del st.session_state["pending_duplicate_info"]
                st.rerun()
        st.stop()
    # ───────────────────────────────────────────────────────────────────────

    # Botones Analyze y New Analysis en la misma fila
    col1, col2 = st.columns([2, 1])
    with col1:
        analyze_clicked = st.button("Analyze Documents", type="primary", use_container_width=True)
    with col2:
        if st.button("🔄 New Analysis", type="secondary", use_container_width=True):
            st.session_state.show_confirm = True
            st.rerun()

    # Confirmación antes de limpiar
    if st.session_state.get("show_confirm", False):
        st.warning("⚠️ This will clear the current analysis. Are you sure?")
        col_yes, col_no = st.columns([1, 1])
        with col_yes:
            if st.button("✅ Yes, run new analysis", type="primary",
                         key="btn_confirm_new", use_container_width=True):
                for key in list(st.session_state.keys()):
                    if any(word in key.lower() for word in
                           ["decision", "cambio", "riesgo", "analys", "result", "finding", "confirm"]):
                        del st.session_state[key]
                st.session_state.auto_run_analysis = True
                st.rerun()
        with col_no:
            if st.button("❌ Cancel", key="btn_cancel_new", use_container_width=True):
                st.session_state.show_confirm = False
                st.rerun()

    should_run_analysis = analyze_clicked or st.session_state.get("auto_run_analysis", False)

    if should_run_analysis:
        is_auto_run = st.session_state.get("auto_run_analysis", False)

        if not uploaded_files:
            if is_auto_run:
                st.warning("⚠️ Please upload documents first.")
                del st.session_state["auto_run_analysis"]
            else:
                st.warning("⚠️ Please upload at least one document before analyzing.")
        elif not st.session_state.project_name.strip():
            if is_auto_run:
                del st.session_state["auto_run_analysis"]
            st.warning("⚠️ Please enter a project name before analyzing.")
            st.stop()
        else:
            api_key = os.getenv("ANTHROPIC_API_KEY", "").strip()

            if not api_key:
                if is_auto_run:
                    del st.session_state["auto_run_analysis"]
                st.error("❌ API Key not found. Please add ANTHROPIC_API_KEY to your environment variables.")
            else:
                has_cached_results = (
                    not is_auto_run
                    and st.session_state.decisiones is not None
                    and st.session_state.findings_by_category is not None
                    and current_signature == st.session_state.last_analyzed_signature
                )

                if not has_cached_results:
                    raw_text = extract_all_text(uploaded_files)
                    char_count = len(raw_text)

                    if char_count > MAX_CHARS:
                        raw_text = raw_text[:MAX_CHARS]
                        char_count = MAX_CHARS
                        st.warning(
                            "⚠️ Documents truncated to 50,000 characters due to context "
                            "limits -- consider uploading fewer files per analysis."
                        )

                    num_docs = len(uploaded_files)
                    document_names = [f.name for f in uploaded_files]
                    project_name = st.session_state.project_name.strip()
                    project_description = st.session_state.project_description.strip()

                    # Calcular hash del conjunto de documentos
                    doc_hash = calculate_documents_hash(uploaded_files)

                    # Verificar si estos documentos ya fueron analizados
                    project_id_check = get_or_create_project(supabase_client, project_name, project_description)
                    duplicate = check_duplicate_analysis(supabase_client, project_id_check, doc_hash)

                    # Si hay duplicado y el usuario no confirmó re-análisis, mostrar aviso
                    if duplicate and not st.session_state.get("force_reanalysis"):
                        prev_date = format_analysis_date(duplicate.get("created_at", ""))
                        prev_docs = ", ".join(duplicate.get("documents_analyzed", []))
                        # Guardar info del duplicado en session_state para sobrevivir el rerun
                        st.session_state.pending_duplicate_info = f"⚠️ These exact documents were already analyzed on **{prev_date}** ({prev_docs}). Running again will create a new record with the same content."
                        st.rerun()

                    # Limpiar flag de re-análisis forzado después de usarlo
                    if st.session_state.get("force_reanalysis"):
                        del st.session_state["force_reanalysis"]

                    try:
                        # Cargar contexto del proyecto para enriquecer el análisis
                        project_ctx = {}
                        if project_id_check:
                            project_ctx = load_project_context(supabase_client, project_id_check)

                        with st.spinner("🔍 Stellae is scanning your documents -- this may take 20-40 seconds..."):
                            response_text = run_anthropic_analysis(raw_text, project_context=project_ctx)

                        parse_and_store_results(response_text, num_docs, char_count)
                        st.session_state.last_analyzed_signature = current_signature

                        project_id = project_id_check
                        st.session_state.project_id = project_id

                        findings_by_category = run_governance_pipeline(supabase_client, project_id)

                        try:
                            save_analysis_to_supabase(
                                supabase_client, project_id, document_names,
                                char_count, response_text, findings_by_category,
                                doc_hash=doc_hash,
                            )
                            st.success("✅ Analysis saved to Stellae database -- Project: " + project_name)
                        except Exception as save_error:
                            st.warning("⚠️ Analysis complete but could not save to database: " + str(save_error))

                    except Exception as e:
                        st.error("❌ Analysis failed: " + str(e))

                    if is_auto_run and "auto_run_analysis" in st.session_state:
                        del st.session_state["auto_run_analysis"]
                elif is_auto_run and "auto_run_analysis" in st.session_state:
                    del st.session_state["auto_run_analysis"]

    # Mostrar resultados si existen — con o sin archivos en el uploader
    # La firma solo se verifica si hay archivos cargados actualmente
    has_results = (
        st.session_state.get("decisiones") is not None
        and st.session_state.get("findings_by_category") is not None
    )
    files_match = (
        not uploaded_files  # Sin archivos — mostrar últimos resultados siempre
        or current_signature == st.session_state.get("last_analyzed_signature")
    )
    if has_results and files_match:
        render_analysis_results_tabs()


def render_governance_page(supabase_client: Client) -> None:
    """Página Governance: configuración completa de gobernanza por proyecto."""
    st.title("🏛️ Governance Setup")
    render_project_badge(st.session_state.get("project_name", ""))

    project_id = st.session_state.get("project_id")
    selected_name = st.session_state.get("project_name", "")

    if not project_id:
        st.info("ℹ️ Select a project in the sidebar to configure governance.")
        return

    # Cargar datos del proyecto
    proj_data = load_project_context(supabase_client, project_id)

    # =========================================================================
    # 1. PROJECT CONTEXT — expander
    # =========================================================================
    _ctx_expanded = st.session_state.get("ctx_expander_open", False)
    # Mostrar resumen del contexto guardado debajo del expander title
    _ctx_label = "📋 Project Context"
    if proj_data.get("context_filename") or proj_data.get("industry"):
        _ctx_parts = []
        if proj_data.get("industry"): _ctx_parts.append(proj_data["industry"])
        if proj_data.get("project_stage"): _ctx_parts.append(proj_data["project_stage"])
        if proj_data.get("context_filename"): _ctx_parts.append(f"📄 {proj_data['context_filename']}")
        _ctx_label = f"📋 Project Context ✅ — {' · '.join(_ctx_parts)}"
    with st.expander(_ctx_label, expanded=_ctx_expanded):
        # ── Estado actual del contexto — siempre visible ──────────────────
        has_context = bool(proj_data.get("context_filename") or proj_data.get("industry"))
        if has_context:
            # Construir HTML por partes — evita problemas con comillas anidadas en f-string
            _ref_line = ""
            _meta_line = ""
            if proj_data.get("context_filename"):
                _ref_line = f"<div style='color:#e7e9ea; font-size:13px; margin-bottom:4px;'>📄 <b>Ref doc:</b> {proj_data['context_filename']}</div>"
            _meta_parts = [x for x in [proj_data.get("industry",""), proj_data.get("project_type",""), proj_data.get("project_stage","")] if x]
            if _meta_parts:
                _meta_line = f"<div style='color:#8899a6; font-size:12px;'>{' · '.join(_meta_parts)}</div>"
            st.markdown(
                f"""<div style="background:rgba(30,58,95,0.4);border:1px solid rgba(201,168,76,0.4);
                    border-left:3px solid #C9A84C;border-radius:8px;padding:10px 16px;margin-bottom:16px;">
                    <div style="color:#C9A84C;font-size:11px;letter-spacing:1px;margin-bottom:6px;">✅ CONTEXT CONFIGURED</div>
                    {_ref_line}{_meta_line}
                </div>""",
                unsafe_allow_html=True
            )
        else:
            st.info("ℹ️ No context configured yet. Fill in the fields below and save.", icon="📋")

        # ── Reference Document uploader ─────────────────────────────────────
        st.markdown("**Reference Document** *(optional — service design, procedures, project charter)*")
        if proj_data.get("context_filename"):
            st.caption(f"📄 Loaded: **{proj_data['context_filename']}** — upload a new file to replace it.")

        context_file = st.file_uploader(
            "Upload reference document",
            type=["pdf", "docx", "txt"],
            key="context_doc_uploader",
            label_visibility="collapsed",
        )

        st.markdown("---")

        # ── Campos de texto — pre-cargar desde Supabase en session_state ──
        # Usar project_id en la key garantiza reset al cambiar de proyecto
        _sk_ind  = f"ctx_ind_{project_id}"
        _sk_type = f"ctx_type_{project_id}"
        _sk_stg  = f"ctx_stg_{project_id}"
        _sk_desc = f"ctx_desc_{project_id}"
        if _sk_ind not in st.session_state:
            st.session_state[_sk_ind]  = proj_data.get("industry") or ""
        if _sk_type not in st.session_state:
            st.session_state[_sk_type] = proj_data.get("project_type") or ""
        if _sk_stg not in st.session_state:
            st.session_state[_sk_stg]  = proj_data.get("project_stage") or ""
        if _sk_desc not in st.session_state:
            st.session_state[_sk_desc] = proj_data.get("description") or ""

        col1, col2 = st.columns(2)
        with col1:
            industry = st.text_input(
                "Industry / Sector",
                placeholder="Ex: Oil & Gas, Infrastructure, Mining, Defense",
                key=_sk_ind,
            )
            project_type = st.text_input(
                "Project Type",
                placeholder="Ex: EPC, EPCM, Design-Build, O&M",
                key=_sk_type,
            )
        with col2:
            project_stage = st.text_input(
                "Current Stage",
                placeholder="Ex: FEED, Detailed Engineering, Construction, Commissioning",
                key=_sk_stg,
            )
            description = st.text_area(
                "Project Description",
                placeholder="Brief description of the project objectives and scope",
                height=80,
                key=_sk_desc,
            )

        if st.button("💾 Save Project Context", type="primary", key="btn_save_ctx"):
            context_text = proj_data.get("context_document")
            context_filename = proj_data.get("context_filename")
            if context_file:
                ext = context_file.name.rsplit(".", 1)[-1].lower()
                if ext == "pdf":
                    context_text = extract_text_from_pdf(context_file)
                elif ext == "docx":
                    context_text = extract_text_from_docx(context_file)
                elif ext == "txt":
                    context_text = extract_text_from_txt(context_file)
                context_filename = context_file.name
            try:
                supabase_client.table("projects").update({
                    "description": description,
                    "industry": industry,
                    "project_type": project_type,
                    "project_stage": project_stage,
                    "context_document": context_text,
                    "context_filename": context_filename,
                }).eq("id", project_id).execute()
                st.session_state.ctx_expander_open = True
                # Actualizar session_state con valores recién guardados
                st.session_state[f"ctx_ind_{project_id}"]  = industry
                st.session_state[f"ctx_type_{project_id}"] = project_type
                st.session_state[f"ctx_stg_{project_id}"]  = project_stage
                st.session_state[f"ctx_desc_{project_id}"] = description
                # Mostrar confirmacion con los datos guardados
                st.success(
                    f"✅ Project context saved — "
                    f"{'Industry: ' + industry + ' · ' if industry else ''}"
                    f"{'Type: ' + project_type + ' · ' if project_type else ''}"
                    f"{'Stage: ' + project_stage if project_stage else ''}"
                    f"{'· Ref doc: ' + context_filename if context_filename else ''}"
                )
                # Recargar proj_data sin rerun para que los campos muestren valores
                proj_data = load_project_context(supabase_client, project_id)
            except Exception as e:
                st.error(f"❌ Failed to save context: {e}")



    # =========================================================================
    # 2. GOVERNANCE DEFINITIONS
    # =========================================================================
    st.markdown("### Governance Definitions")

    # ── 2A. AI-Assisted Generator ──────────────────────────────────────────
    with st.expander("🤖 AI-Assisted Governance Generator", expanded=False):
        st.markdown(
            "Upload a project document and Stellae will automatically extract roles "
            "and escalation rules. Review the proposal and activate with one click."
        )

        gen_tab_a, gen_tab_b = st.tabs([
            "📎 Use reference document already uploaded",
            "📤 Upload a new document"
        ])

        use_existing = False
        use_new = False
        gen_file = None

        with gen_tab_a:
            if proj_data.get("context_document") and proj_data.get("context_filename"):
                st.success(f"✅ Reference document available: **{proj_data['context_filename']}**")
                st.caption("Stellae will read this document to extract governance structure.")
                use_existing = st.button(
                    "🤖 Generate governance from reference document",
                    key="gen_from_existing",
                    type="primary",
                    use_container_width=True
                )
            else:
                st.info("ℹ️ No reference document uploaded yet. Save one in **Project Context** above first.")

        with gen_tab_b:
            st.caption("PEP, Project Charter, RACI matrix, contract, Delegation of Authority.")
            gen_file = st.file_uploader(
                "Choose file",
                type=["pdf", "docx", "txt"],
                key="governance_gen_uploader",
                label_visibility="collapsed"
            )
            if gen_file:
                st.success(f"✅ Ready: **{gen_file.name}**")
                use_new = st.button(
                    "🤖 Generate governance from this document",
                    key="gen_from_new",
                    type="primary",
                    use_container_width=True
                )

        # Ejecutar generación
        gen_source_text = None
        gen_source_name = None

        if use_existing and proj_data.get("context_document"):
            gen_source_text = proj_data["context_document"]
            gen_source_name = proj_data.get("context_filename", "reference document")
        elif use_new and gen_file:
            ext = gen_file.name.rsplit(".", 1)[-1].lower()
            if ext == "pdf":
                gen_source_text = extract_text_from_pdf(gen_file)
            elif ext == "docx":
                gen_source_text = extract_text_from_docx(gen_file)
            else:
                gen_source_text = extract_text_from_txt(gen_file)
            gen_source_name = gen_file.name

        if gen_source_text:
            with st.spinner(f"🤖 Stellae is reading **{gen_source_name}** and extracting governance structure..."):
                try:
                    import json
                    suggestion = auto_generate_governance(gen_source_text, selected_name)
                    st.session_state.governance_suggestion = suggestion
                    st.session_state.governance_suggestion_source = gen_source_name
                    st.rerun()
                except json.JSONDecodeError as e:
                    st.error(f"❌ Could not parse AI response. Try with a different document. Error: {e}")
                except Exception as e:
                    st.error(f"❌ Generation failed: {e}")

    # ── 2B. Manual Input ───────────────────────────────────────────────────
    with st.expander("✏️ Manual Roles & Rules Input", expanded=False):

        # Import desde archivo
        imp_tab_roles, imp_tab_rules = st.tabs(["📂 Import Roles from file", "📂 Import Rules from file"])

        with imp_tab_roles:
            st.caption("Format: Role Name | Authority Level (1-5) | Can Approve (Yes/No) | Max Impact (USD)")
            roles_file = st.file_uploader("Upload document with roles", type=["pdf", "docx", "txt"],
                                           key="roles_import_uploader")
            if roles_file and st.button("🔍 Parse and import roles", key="parse_roles_btn"):
                ext = roles_file.name.rsplit(".", 1)[-1].lower()
                if ext == "pdf":
                    raw_text = extract_text_from_pdf(roles_file)
                elif ext == "docx":
                    raw_text = extract_text_from_docx(roles_file)
                else:
                    raw_text = extract_text_from_txt(roles_file)
                parsed_roles = parse_roles_from_text(raw_text)
                if parsed_roles:
                    st.success(f"Found {len(parsed_roles)} role(s). Review and confirm:")
                    for i, role in enumerate(parsed_roles):
                        st.markdown(f"**{role['role_name']}** — Level {role['authority_level']} — {'✓ Approves' if role['can_approve_changes'] else '✗ No approval'} — Max: ${role['max_impact_value']:,.0f}")
                    if st.button("✅ Import all roles", key="confirm_import_roles"):
                        imported = 0
                        for role in parsed_roles:
                            try:
                                supabase_client.table("governance_roles").insert({"project_id": project_id, **role}).execute()
                                imported += 1
                            except Exception:
                                pass
                        st.success(f"✅ {imported} role(s) imported successfully.")
                        st.rerun()
                else:
                    st.warning("No roles detected. Check your document format.")

        with imp_tab_rules:
            st.caption("Format: Rule Name | Category (decision/change/risk) | Keywords (comma-separated) | Required Level (1-5)")
            rules_file = st.file_uploader("Upload document with rules", type=["pdf", "docx", "txt"],
                                           key="rules_import_uploader")
            if rules_file and st.button("🔍 Parse and import rules", key="parse_rules_btn"):
                ext = rules_file.name.rsplit(".", 1)[-1].lower()
                if ext == "pdf":
                    raw_text = extract_text_from_pdf(rules_file)
                elif ext == "docx":
                    raw_text = extract_text_from_docx(rules_file)
                else:
                    raw_text = extract_text_from_txt(rules_file)
                parsed_rules = parse_rules_from_text(raw_text)
                if parsed_rules:
                    st.success(f"Found {len(parsed_rules)} rule(s). Review and confirm:")
                    for rule in parsed_rules:
                        st.markdown(f"**{rule['rule_name']}** — {rule['category']} — Keywords: {', '.join(rule['trigger_keywords'])} — Level {rule['required_authority_level']} required")
                    if st.button("✅ Import all rules", key="confirm_import_rules"):
                        imported = 0
                        for rule in parsed_rules:
                            try:
                                supabase_client.table("escalation_rules").insert({"project_id": project_id, **rule}).execute()
                                imported += 1
                            except Exception:
                                pass
                        st.success(f"✅ {imported} rule(s) imported successfully.")
                        st.rerun()
                else:
                    st.warning("No rules detected. Check your document format.")

        st.divider()

        # Formulario manual de roles
        st.markdown("**Add role manually**")
        with st.form("add_role_form", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                role_name = st.text_input("Role name", placeholder="Ex: Project Director, PMO, Steering Committee")
                authority_level = st.selectbox(
                    "Authority level",
                    options=list(AUTHORITY_LEVEL_LABELS.keys()),
                    format_func=lambda x: AUTHORITY_LEVEL_LABELS[x],
                )
            with col2:
                can_approve_changes = st.checkbox("Can approve changes")
                max_impact_value = st.number_input("Max approvable impact (USD)", min_value=0.0, value=0.0, step=1000.0)
            add_role = st.form_submit_button("➕ Add Role", type="primary")

        if add_role:
            if not role_name.strip():
                st.warning("Please enter a role name.")
            else:
                try:
                    supabase_client.table("governance_roles").insert({
                        "project_id": project_id,
                        "role_name": role_name.strip(),
                        "authority_level": authority_level,
                        "can_approve_changes": can_approve_changes,
                        "max_impact_value": max_impact_value,
                    }).execute()
                    st.success(f"Role '{role_name}' added.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Failed to add role: {e}")

        st.divider()

        # Formulario manual de rules
        st.markdown("**Add escalation rule manually**")
        with st.form("add_rule_form", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                rule_name = st.text_input("Rule name", placeholder="Ex: Fire safety changes require SC approval")
                category = st.selectbox("Category", options=["decision", "change", "risk"])
                is_active = st.checkbox("Rule active", value=True)
            with col2:
                trigger_keywords_raw = st.text_input("Trigger keywords", placeholder="comma-separated: fire, safety, permit")
                required_authority_level = st.selectbox(
                    "Required authority level",
                    options=list(AUTHORITY_LEVEL_LABELS.keys()),
                    format_func=lambda x: AUTHORITY_LEVEL_LABELS[x],
                )
            add_rule = st.form_submit_button("➕ Add Rule", type="primary")

        if add_rule:
            if not rule_name.strip():
                st.warning("Please enter a rule name.")
            else:
                keywords_list = [k.strip() for k in trigger_keywords_raw.split(",") if k.strip()]
                try:
                    supabase_client.table("escalation_rules").insert({
                        "project_id": project_id,
                        "rule_name": rule_name.strip(),
                        "category": category,
                        "trigger_keywords": keywords_list,
                        "required_authority_level": required_authority_level,
                        "is_active": is_active,
                    }).execute()
                    st.success(f"Rule '{rule_name}' added.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Failed to add rule: {e}")

        st.caption("💡 Define keywords in the same language as your project documents.")

    # =========================================================================
    # 3. PROPUESTA AI — fuera de expanders, visible tras generación
    # =========================================================================
    if st.session_state.get("governance_suggestion"):
        suggestion = st.session_state.governance_suggestion
        source = st.session_state.get("governance_suggestion_source", "document")

        st.divider()
        st.success(f"✅ Stellae analyzed **{source}** and found the following governance structure:")
        st.markdown("**Review and activate what you want to use:**")

        suggested_roles = suggestion.get("roles", [])
        roles_to_import = []
        if suggested_roles:
            st.markdown(f"**🏛️ Suggested Roles ({len(suggested_roles)})**")
            for i, role in enumerate(suggested_roles):
                col_check, col_info = st.columns([1, 8])
                with col_check:
                    include = st.checkbox("", value=True, key=f"include_role_{i}")
                with col_info:
                    approve_text = "✓ Can approve" if role.get("can_approve_changes") else "✗ No approval"
                    max_text = f"Max: ${role.get('max_impact_value', 0):,.0f}" if role.get("max_impact_value", 0) > 0 else "No limit"
                    st.markdown(f"**{role.get('role_name')}** — Level {role.get('authority_level')} — {approve_text} — {max_text}")
                if include:
                    roles_to_import.append(role)

        suggested_rules = suggestion.get("rules", [])
        rules_to_import = []
        if suggested_rules:
            st.markdown(f"**📋 Suggested Escalation Rules ({len(suggested_rules)})**")
            for i, rule in enumerate(suggested_rules):
                col_check, col_info = st.columns([1, 8])
                with col_check:
                    include = st.checkbox("", value=True, key=f"include_rule_{i}")
                with col_info:
                    keywords_str = ", ".join(rule.get("trigger_keywords", []))
                    st.markdown(f"**{rule.get('rule_name')}** — {rule.get('category')} — Level {rule.get('required_authority_level')} required")
                    st.caption(f"Keywords: {keywords_str}")
                if include:
                    rules_to_import.append(rule)

        st.divider()
        col_confirm, col_cancel = st.columns([2, 1])
        with col_confirm:
            if st.button("✅ Activate selected roles and rules", type="primary", use_container_width=True):
                imported_roles = 0
                imported_rules = 0
                for role in roles_to_import:
                    try:
                        supabase_client.table("governance_roles").insert({
                            "project_id": project_id,
                            "role_name": role.get("role_name", "Unknown"),
                            "authority_level": role.get("authority_level", 3),
                            "can_approve_changes": role.get("can_approve_changes", False),
                            "max_impact_value": role.get("max_impact_value", 0),
                        }).execute()
                        imported_roles += 1
                    except Exception:
                        pass
                for rule in rules_to_import:
                    try:
                        supabase_client.table("escalation_rules").insert({
                            "project_id": project_id,
                            "rule_name": rule.get("rule_name", "Unknown"),
                            "category": rule.get("category", "change"),
                            "trigger_keywords": rule.get("trigger_keywords", []),
                            "required_authority_level": rule.get("required_authority_level", 5),
                            "is_active": True,
                        }).execute()
                        imported_rules += 1
                    except Exception:
                        pass
                del st.session_state["governance_suggestion"]
                if "governance_suggestion_source" in st.session_state:
                    del st.session_state["governance_suggestion_source"]
                st.success(f"✅ {imported_roles} role(s) and {imported_rules} rule(s) activated successfully.")
                st.rerun()
        with col_cancel:
            if st.button("❌ Discard proposal", use_container_width=True):
                del st.session_state["governance_suggestion"]
                if "governance_suggestion_source" in st.session_state:
                    del st.session_state["governance_suggestion_source"]
                st.rerun()

    # =========================================================================
    # 4. ACTIVE ROLES — fuera de expanders
    # =========================================================================
    st.divider()
    st.markdown("### Active Roles")

    roles_response = (
        supabase_client.table("governance_roles")
        .select("*")
        .eq("project_id", project_id)
        .execute()
    )
    roles_df = pd.DataFrame(roles_response.data) if roles_response.data else pd.DataFrame()

    if not roles_df.empty:
        for _, role in roles_df.iterrows():
            role_id = role["id"]
            if st.session_state.get("editing_role") == role_id:
                with st.form(key=f"edit_role_form_{role_id}"):
                    st.markdown(f"**Editing: {role['role_name']}**")
                    ec1, ec2 = st.columns(2)
                    with ec1:
                        new_role_name = st.text_input("Role name", value=role["role_name"])
                        new_authority = st.selectbox(
                            "Authority level",
                            options=list(AUTHORITY_LEVEL_LABELS.keys()),
                            index=int(role["authority_level"]) - 1,
                            format_func=lambda x: AUTHORITY_LEVEL_LABELS[x]
                        )
                    with ec2:
                        new_can_approve = st.checkbox("Can approve changes", value=bool(role["can_approve_changes"]))
                        new_max_impact = st.number_input(
                            "Max approvable impact (USD)",
                            min_value=0.0,
                            value=float(role["max_impact_value"]) if role["max_impact_value"] else 0.0,
                            step=1000.0
                        )
                    col_save, col_cancel_edit = st.columns(2)
                    with col_save:
                        save_edit = st.form_submit_button("💾 Save changes", type="primary", use_container_width=True)
                    with col_cancel_edit:
                        cancel_edit = st.form_submit_button("❌ Cancel", use_container_width=True)
                if save_edit:
                    supabase_client.table("governance_roles").update({
                        "role_name": new_role_name,
                        "authority_level": new_authority,
                        "can_approve_changes": new_can_approve,
                        "max_impact_value": new_max_impact,
                    }).eq("id", role_id).execute()
                    del st.session_state["editing_role"]
                    st.toast("✅ Role updated")
                    st.rerun()
                if cancel_edit:
                    del st.session_state["editing_role"]
                    st.rerun()
            else:
                col1, col2, col3, col4, col5, col6 = st.columns([2, 1, 1, 2, 1, 1])
                with col1:
                    st.text(role["role_name"])
                with col2:
                    st.text(f"Level {role['authority_level']}")
                with col3:
                    st.text("✓ Approves" if role["can_approve_changes"] else "✗ No approval")
                with col4:
                    st.text(f"Max: ${role['max_impact_value']:,.0f}" if role["max_impact_value"] else "No limit")
                with col5:
                    if st.button("✏️", key=f"edit_role_{role_id}", help="Edit"):
                        st.session_state.editing_role = role_id
                        st.rerun()
                with col6:
                    if st.button("🗑️", key=f"del_role_{role_id}"):
                        st.session_state.confirm_delete_role = role_id
                        st.session_state.confirm_delete_role_name = role["role_name"]
                        st.rerun()
    else:
        st.info("No roles defined yet. Use the generator or manual input above.")

    if st.session_state.get("confirm_delete_role"):
        st.warning(f"⚠️ Delete role '{st.session_state.confirm_delete_role_name}'? This cannot be undone.")
        col_yes, col_no = st.columns([1, 1])
        with col_yes:
            if st.button("✅ Yes, delete", key="confirm_yes_role", type="primary", use_container_width=True):
                supabase_client.table("governance_roles").delete().eq("id", st.session_state.confirm_delete_role).execute()
                del st.session_state["confirm_delete_role"]
                del st.session_state["confirm_delete_role_name"]
                st.toast("✅ Role deleted")
                st.rerun()
        with col_no:
            if st.button("❌ Cancel", key="confirm_no_role", use_container_width=True):
                del st.session_state["confirm_delete_role"]
                del st.session_state["confirm_delete_role_name"]
                st.rerun()

    # =========================================================================
    # 5. ACTIVE ESCALATION RULES — fuera de expanders
    # =========================================================================
    st.divider()
    st.markdown("### Active Escalation Rules")
    st.caption("💡 Define keywords in the same language as your project documents.")

    rules_response = (
        supabase_client.table("escalation_rules")
        .select("*")
        .eq("project_id", project_id)
        .execute()
    )
    rules_df = pd.DataFrame(rules_response.data) if rules_response.data else pd.DataFrame()

    if not rules_df.empty:
        for _, rule in rules_df.iterrows():
            rule_id = rule["id"]
            if st.session_state.get("editing_rule") == rule_id:
                with st.form(key=f"edit_rule_form_{rule_id}"):
                    st.markdown(f"**Editing: {rule['rule_name']}**")
                    er1, er2 = st.columns(2)
                    with er1:
                        new_rule_name = st.text_input("Rule name", value=rule["rule_name"])
                        new_category = st.selectbox(
                            "Category",
                            options=["decision", "change", "risk"],
                            index=["decision", "change", "risk"].index(rule["category"])
                            if rule["category"] in ["decision", "change", "risk"] else 1
                        )
                        new_is_active = st.checkbox("Rule active", value=bool(rule["is_active"]))
                    with er2:
                        current_keywords = ", ".join(rule["trigger_keywords"]) if rule["trigger_keywords"] else ""
                        new_keywords_raw = st.text_area(
                            "Keywords (comma-separated)",
                            value=current_keywords,
                            height=100,
                        )
                        new_required_level = st.selectbox(
                            "Required authority level",
                            options=list(AUTHORITY_LEVEL_LABELS.keys()),
                            index=int(rule["required_authority_level"]) - 1,
                            format_func=lambda x: AUTHORITY_LEVEL_LABELS[x]
                        )
                    col_save_r, col_cancel_r = st.columns(2)
                    with col_save_r:
                        save_rule_edit = st.form_submit_button("💾 Save changes", type="primary", use_container_width=True)
                    with col_cancel_r:
                        cancel_rule_edit = st.form_submit_button("❌ Cancel", use_container_width=True)
                if save_rule_edit:
                    new_keywords_list = [k.strip() for k in new_keywords_raw.split(",") if k.strip()]
                    supabase_client.table("escalation_rules").update({
                        "rule_name": new_rule_name,
                        "category": new_category,
                        "trigger_keywords": new_keywords_list,
                        "required_authority_level": new_required_level,
                        "is_active": new_is_active,
                    }).eq("id", rule_id).execute()
                    del st.session_state["editing_rule"]
                    st.toast("✅ Rule updated")
                    st.rerun()
                if cancel_rule_edit:
                    del st.session_state["editing_rule"]
                    st.rerun()
            else:
                col1, col2, col3, col4, col5, col6 = st.columns([2, 1, 3, 1, 1, 1])
                with col1:
                    st.text(rule["rule_name"])
                with col2:
                    st.text(rule["category"])
                with col3:
                    keywords = ", ".join(rule["trigger_keywords"]) if rule["trigger_keywords"] else "None"
                    if len(keywords) > 40:
                        st.caption(f"Keywords: {keywords}")
                    else:
                        st.text(f"Keywords: {keywords}")
                with col4:
                    st.text(f"Level {rule['required_authority_level']}")
                with col5:
                    if st.button("✏️", key=f"edit_rule_{rule_id}", help="Edit"):
                        st.session_state.editing_rule = rule_id
                        st.rerun()
                with col6:
                    if st.button("🗑️", key=f"del_rule_{rule_id}"):
                        st.session_state.confirm_delete_rule = rule_id
                        st.session_state.confirm_delete_rule_name = rule["rule_name"]
                        st.rerun()
    else:
        st.info("No escalation rules defined yet. Use the generator or manual input above.")

    if st.session_state.get("confirm_delete_rule"):
        st.warning(f"⚠️ Delete rule '{st.session_state.confirm_delete_rule_name}'? This cannot be undone.")
        col_yes, col_no = st.columns([1, 1])
        with col_yes:
            if st.button("✅ Yes, delete", key="confirm_yes_rule", type="primary", use_container_width=True):
                supabase_client.table("escalation_rules").delete().eq("id", st.session_state.confirm_delete_rule).execute()
                del st.session_state["confirm_delete_rule"]
                del st.session_state["confirm_delete_rule_name"]
                st.toast("✅ Rule deleted")
                st.rerun()
        with col_no:
            if st.button("❌ Cancel", key="confirm_no_rule", use_container_width=True):
                del st.session_state["confirm_delete_rule"]
                del st.session_state["confirm_delete_rule_name"]
                st.rerun()





def render_audit_trail_page(supabase_client: Client) -> None:
    """Página Audit Trail: historial de análisis y edición de status."""
    st.title("📋 Audit Trail")
    render_project_badge(st.session_state.get("project_name", ""))
    st.markdown(
        '<p class="subtitle">Review past analyses and update finding status</p>',
        unsafe_allow_html=True,
    )

    selected_project_id = st.session_state.get("project_id")
    selected_name = st.session_state.get("project_name", "")

    if not selected_project_id:
        st.info("ℹ️ Select a project in the sidebar to view the audit trail.")
        return

    analyses_response = (
        supabase_client.table("analyses")
        .select("*")
        .eq("project_id", selected_project_id)
        .order("created_at", desc=True)
        .execute()
    )

    if not analyses_response.data:
        st.caption("No analyses found for this project.")
        return

    for analysis in analyses_response.data:
        date_label = format_analysis_date(analysis.get("created_at", ""))
        docs = analysis.get("documents_analyzed", [])
        chars = analysis.get("characters_processed", 0)
        analysis_id = analysis["id"]

        with st.expander(f"📅 {date_label} -- {len(docs)} document(s)"):
            # Fila superior: info + botones export + eliminar
            col_info, col_export, col_delete = st.columns([4, 2, 1])
            with col_info:
                st.markdown("**Documents analyzed:**")
                for doc_name in docs:
                    st.markdown(f"- {doc_name}")
                st.markdown(f"**Characters processed:** {chars:,}")
            with col_export:
                st.markdown("&nbsp;", unsafe_allow_html=True)
                # Selector de idioma + botón export para este análisis
                _audit_lang_options = list(OUTPUT_LANGUAGES.keys())
                _audit_lang_key = f"audit_lang_{analysis_id}"
                _audit_lang_sel = st.selectbox(
                    "Language",
                    options=_audit_lang_options,
                    index=_audit_lang_options.index(
                        st.session_state.get("pdf_lang_selected", "English")
                    ),
                    key=_audit_lang_key,
                    label_visibility="collapsed",
                )
                if st.button("📄 Export PDF", key=f"export_analysis_{analysis_id}",
                             use_container_width=True):
                    # Cargar findings de este análisis específico
                    _findings_resp = supabase_client.table("findings").select("*").eq(
                        "analysis_id", analysis_id
                    ).execute()
                    _analysis_findings = _findings_resp.data if _findings_resp.data else []

                    if _analysis_findings:
                        _pdf_lang = OUTPUT_LANGUAGES[_audit_lang_sel]
                        with st.spinner(f"Generating PDF..."):
                            try:
                                _pdf_findings = translate_findings_for_pdf(
                                    _analysis_findings, _pdf_lang
                                )
                                # Status para este análisis individual
                                _open   = sum(1 for f in _analysis_findings if f.get("status") == "open")
                                _review = sum(1 for f in _analysis_findings if f.get("status") == "in_review")
                                _closed = sum(1 for f in _analysis_findings if f.get("status") == "closed")
                                _violations = sum(1 for f in _analysis_findings if f.get("governance_violation"))
                                if _violations > 0:
                                    _slabel, _smsg = "CRITICAL", f"{_violations} governance violation(s) detected"
                                elif _open > 0:
                                    _slabel, _smsg = "AT RISK", f"{_open} open finding(s)"
                                else:
                                    _slabel, _smsg = "UNDER CONTROL", "No critical issues detected"

                                _pdf_bytes = generate_executive_pdf(
                                    selected_name,
                                    _slabel, _smsg,
                                    _pdf_findings,
                                    [analysis],  # solo este análisis
                                    language_code=_pdf_lang,
                                )
                                _fname = (
                                    f"stellae_analysis_{date_label.replace(' ', '_').replace('/', '-')}"
                                    f"_{selected_name.replace(' ', '_')}.pdf"
                                )
                                st.download_button(
                                    label="⬇️ Download PDF",
                                    data=_pdf_bytes,
                                    file_name=_fname,
                                    mime="application/pdf",
                                    key=f"dl_analysis_{analysis_id}",
                                    use_container_width=True,
                                )
                            except Exception as e:
                                st.error(f"❌ PDF error: {e}")
                    else:
                        st.warning("No findings found for this analysis.")

            with col_delete:
                st.markdown("&nbsp;", unsafe_allow_html=True)
                if st.button("🗑️", key=f"del_analysis_{analysis_id}", type="secondary",
                             use_container_width=True, help="Delete this analysis"):
                    st.session_state.confirm_delete_analysis = analysis_id
                    st.session_state.confirm_delete_analysis_date = date_label
                    st.rerun()

            # Confirmación de eliminación
            if st.session_state.get("confirm_delete_analysis") == analysis_id:
                st.warning(
                    f"⚠️ Delete analysis from **{date_label}**? "
                    f"This will permanently remove all {len(docs)} document(s) and all their findings. "
                    f"Dashboard metrics will update automatically."
                )
                col_yes, col_no = st.columns([1, 1])
                with col_yes:
                    if st.button("✅ Yes, delete permanently", key=f"confirm_del_{analysis_id}",
                                 type="primary", use_container_width=True):
                        try:
                            # CASCADE en Supabase elimina findings automáticamente
                            supabase_client.table("analyses").delete().eq("id", analysis_id).execute()
                            del st.session_state["confirm_delete_analysis"]
                            del st.session_state["confirm_delete_analysis_date"]
                            st.toast("✅ Analysis deleted — metrics updated.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ Failed to delete: {e}")
                with col_no:
                    if st.button("❌ Cancel", key=f"cancel_del_{analysis_id}", use_container_width=True):
                        del st.session_state["confirm_delete_analysis"]
                        del st.session_state["confirm_delete_analysis_date"]
                        st.rerun()

            st.markdown("---")
            st.markdown("**Findings -- edit status below:**")
            render_findings_editor(supabase_client, analysis_id)


# =============================================================================
# CONFIGURACIÓN DE PÁGINA STREAMLIT
# =============================================================================

# Favicon embebido en base64 — funciona en local y Railway sin archivos externos
_FAVICON_B64 = "data:image/svg+xml;base64,PHN2ZyB2aWV3Qm94PSIwIDAgNjQgNjQiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyIgd2lkdGg9IjY0IiBoZWlnaHQ9IjY0Ij4KICA8IS0tIEZvbmRvIG5hdnkgcmVkb25kZWFkbyAtLT4KICA8cmVjdCB3aWR0aD0iNjQiIGhlaWdodD0iNjQiIGZpbGw9IiMxRTNBNUYiIHJ4PSIxMiIvPgoKICA8IS0tIEVzdHJlbGxhIGRlIDYgcHVudGFzIGVuIGdvbGQgLS0+CiAgPHBvbHlnb24gcG9pbnRzPSIzMiw4IDM1LDI0IDUwLDI0IDM5LDMzIDQzLDQ5IDMyLDQxIDIxLDQ5IDI1LDMzIDE0LDI0IDI5LDI0IgogICAgZmlsbD0ibm9uZSIgc3Ryb2tlPSIjQzlBODRDIiBzdHJva2Utd2lkdGg9IjIiIHN0cm9rZS1saW5lam9pbj0icm91bmQiLz4KCiAgPCEtLSBOb2RvcyBwcmluY2lwYWxlcyAtLT4KICA8Y2lyY2xlIGN4PSIzMiIgY3k9IjgiICByPSIzIiBmaWxsPSIjQzlBODRDIi8+CiAgPGNpcmNsZSBjeD0iMzIiIGN5PSI1NiIgcj0iMyIgZmlsbD0iI0M5QTg0QyIvPgogIDxjaXJjbGUgY3g9IjUwIiBjeT0iMjQiIHI9IjIuMiIgZmlsbD0iI0M5QTg0QyIgb3BhY2l0eT0iMC43NSIvPgogIDxjaXJjbGUgY3g9IjQzIiBjeT0iNDkiIHI9IjIuMiIgZmlsbD0iI0M5QTg0QyIgb3BhY2l0eT0iMC43NSIvPgogIDxjaXJjbGUgY3g9IjIxIiBjeT0iNDkiIHI9IjIuMiIgZmlsbD0iI0M5QTg0QyIgb3BhY2l0eT0iMC43NSIvPgogIDxjaXJjbGUgY3g9IjE0IiBjeT0iMjQiIHI9IjIuMiIgZmlsbD0iI0M5QTg0QyIgb3BhY2l0eT0iMC43NSIvPgoKICA8IS0tIENlbnRybyAtLT4KICA8Y2lyY2xlIGN4PSIzMiIgY3k9IjMxIiByPSI0IiBmaWxsPSIjQzlBODRDIi8+Cjwvc3ZnPgo="

st.set_page_config(
    page_title="Stellae — Governance Intelligence",
    page_icon=_FAVICON_B64,
    layout="wide",
    initial_sidebar_state="expanded",
)

init_session_state()

if not SUPABASE_URL or not SUPABASE_KEY:
    st.error("❌ Supabase credentials not found. Please add SUPABASE_URL and SUPABASE_KEY to your .env file.")
    st.stop()

supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

# =============================================================================
# AUTENTICACION — pantalla de acceso simple antes de mostrar la app
# La contraseña se configura como variable de entorno APP_PASSWORD en Railway
# =============================================================================
_APP_PASSWORD = os.getenv("APP_PASSWORD", "")

# ── TOKEN DE SESION PERSISTENTE ──────────────────────────────────────────────
# Usar query_params para mantener la sesion aunque Streamlit haga rerun
# El token es un hash del password — no expone el password en la URL
import hashlib as _hashlib

def _make_session_token(password: str) -> str:
    return _hashlib.sha256(password.encode()).hexdigest()[:16]

_SESSION_TOKEN_KEY = "st"  # key corta para la URL

# Verificar si hay token valido en la URL
_url_token = st.query_params.get(_SESSION_TOKEN_KEY, "")
_valid_token = _make_session_token(_APP_PASSWORD) if _APP_PASSWORD else ""

if _url_token and _url_token == _valid_token:
    st.session_state.authenticated = True

if not st.session_state.get("authenticated", False):
    # Centrar el formulario de login
    st.markdown("""
    <style>
    .login-container {
        max-width: 420px;
        margin: 80px auto 0 auto;
        padding: 40px;
        background: #0f1419;
        border: 1px solid #2d3a4f;
        border-radius: 16px;
        border-top: 3px solid #C9A84C;
    }
    </style>
    """, unsafe_allow_html=True)

    col_l, col_c, col_r = st.columns([1, 2, 1])
    with col_c:
        # Logo SVG inline
        st.markdown(
            f'''<div style="text-align:center; margin-bottom:32px;">
                <img src="{_FAVICON_B64}"
                    style="width:72px; height:72px; margin-bottom:16px; display:block; margin-left:auto; margin-right:auto;">
                <div style="color:#e7e9ea; font-size:26px; font-weight:700; letter-spacing:4px; margin-bottom:4px;">STELLAE</div>
                <div style="color:#C9A84C; font-size:12px; letter-spacing:3px; text-transform:uppercase;">Governance Intelligence</div>
            </div>''',
            unsafe_allow_html=True
        )

        with st.form("login_form", clear_on_submit=False):
            pwd = st.text_input(
                "Access code",
                type="password",
                placeholder="Enter your access code",
                label_visibility="collapsed",
            )
            submitted = st.form_submit_button(
                "Enter →", type="primary", use_container_width=True
            )

        if submitted:
            if _APP_PASSWORD and pwd == _APP_PASSWORD:
                st.session_state.authenticated = True
                # Guardar token en URL para persistir la sesion
                st.query_params[_SESSION_TOKEN_KEY] = _make_session_token(_APP_PASSWORD)
                st.rerun()
            elif not _APP_PASSWORD:
                st.error("❌ Access not configured. Contact the administrator.")
            else:
                st.error("❌ Invalid access code.")

        st.markdown(
            '''<div style="text-align:center; margin-top:24px; color:#4a5568; font-size:12px;">
                Request access at stellaeprojects.com
            </div>''',
            unsafe_allow_html=True
        )
    st.stop()
# =============================================================================

st.markdown(
    """
    <style>
    .stApp { background-color: #0f1419; color: #e7e9ea; }
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #1a2332 0%, #0f1419 100%);
        border-right: 1px solid #2d3a4f;
    }
    [data-testid="stSidebar"] .stMarkdown h1 {
        color: #5b9fd4; font-size: 1.75rem; font-weight: 700;
        letter-spacing: 0.05em; margin-bottom: 0.25rem;
    }
    h1, h2, h3 { color: #e7e9ea !important; }
    .subtitle { color: #8899a6; font-size: 1.1rem; margin-bottom: 2rem; }
    [data-testid="stFileUploader"] {
        background-color: #1a2332; border: 1px dashed #3d5266;
        border-radius: 8px; padding: 1rem;
    }
    .stButton > button {
        background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%);
        color: white; border: none; border-radius: 6px;
        padding: 0.6rem 1.5rem; font-weight: 600;
    }
    .file-card {
        background-color: #1a2332; border: 1px solid #2d3a4f;
        border-radius: 6px; padding: 0.75rem 1rem; margin-bottom: 0.5rem;
        display: flex; justify-content: space-between; align-items: center;
    }
    .file-name { color: #e7e9ea; font-weight: 500; }
    .file-size { color: #5b9fd4; font-size: 0.9rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

# =============================================================================
# SIDEBAR -- Navegación multi-página y contexto de proyecto
# =============================================================================

with st.sidebar:
    st.markdown(
        '''<img src="data:image/svg+xml;base64,PHN2ZyB2aWV3Qm94PSIwIDAgMjAwIDQ4IiB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIyMDAiIGhlaWdodD0iNDgiPgogIDxwb2x5Z29uIHBvaW50cz0iMjEuMzIsMTcuNTMgMjQuMDAsNi4wMCAyNC4wMCwyNC4wMCIgZmlsbD0iIzFFM0E1RiIvPgogIDxwb2x5Z29uIHBvaW50cz0iMjQuMDAsNi4wMCAyNi42OCwxNy41MyAyNC4wMCwyNC4wMCIgZmlsbD0iIzhhYTRiYyIvPgogIDxwb2x5Z29uIHBvaW50cz0iMjYuNjgsMTcuNTMgMzYuNzMsMTEuMjcgMjQuMDAsMjQuMDAiIGZpbGw9IiNiMGM0ZDgiLz4KICA8cG9seWdvbiBwb2ludHM9IjM2LjczLDExLjI3IDMwLjQ3LDIxLjMyIDI0LjAwLDI0LjAwIiBmaWxsPSIjMmQ1MDgwIi8+CiAgPHBvbHlnb24gcG9pbnRzPSIzMC40NywyMS4zMiA0Mi4wMCwyNC4wMCAyNC4wMCwyNC4wMCIgZmlsbD0iIzNkNjA5MCIvPgogIDxwb2x5Z29uIHBvaW50cz0iNDIuMDAsMjQuMDAgMzAuNDcsMjYuNjggMjQuMDAsMjQuMDAiIGZpbGw9IiNDOUE4NEMiLz4KICA8cG9seWdvbiBwb2ludHM9IjMwLjQ3LDI2LjY4IDM2LjczLDM2LjczIDI0LjAwLDI0LjAwIiBmaWxsPSIjZDRiNDVhIi8+CiAgPHBvbHlnb24gcG9pbnRzPSIzNi43MywzNi43MyAyNi42OCwzMC40NyAyNC4wMCwyNC4wMCIgZmlsbD0iIzhhNmExYSIvPgogIDxwb2x5Z29uIHBvaW50cz0iMjYuNjgsMzAuNDcgMjQuMDAsNDIuMDAgMjQuMDAsMjQuMDAiIGZpbGw9IiMxRTNBNUYiLz4KICA8cG9seWdvbiBwb2ludHM9IjI0LjAwLDQyLjAwIDIxLjMyLDMwLjQ3IDI0LjAwLDI0LjAwIiBmaWxsPSIjOGFhNGJjIi8+CiAgPHBvbHlnb24gcG9pbnRzPSIyMS4zMiwzMC40NyAxMS4yNywzNi43MyAyNC4wMCwyNC4wMCIgZmlsbD0iI2IwYzRkOCIvPgogIDxwb2x5Z29uIHBvaW50cz0iMTEuMjcsMzYuNzMgMTcuNTMsMjYuNjggMjQuMDAsMjQuMDAiIGZpbGw9IiMyZDUwODAiLz4KICA8cG9seWdvbiBwb2ludHM9IjE3LjUzLDI2LjY4IDYuMDAsMjQuMDAgMjQuMDAsMjQuMDAiIGZpbGw9IiNDOUE4NEMiLz4KICA8cG9seWdvbiBwb2ludHM9IjYuMDAsMjQuMDAgMTcuNTMsMjEuMzIgMjQuMDAsMjQuMDAiIGZpbGw9IiMxRTNBNUYiLz4KICA8cG9seWdvbiBwb2ludHM9IjE3LjUzLDIxLjMyIDExLjI3LDExLjI3IDI0LjAwLDI0LjAwIiBmaWxsPSIjMmQ1MDgwIi8+CiAgPHBvbHlnb24gcG9pbnRzPSIxMS4yNywxMS4yNyAyMS4zMiwxNy41MyAyNC4wMCwyNC4wMCIgZmlsbD0iIzhhYTRiYyIvPgoKICA8IS0tIFNFUEFSQURPUiAtLT4KICA8bGluZSB4MT0iNTQiIHkxPSI4IiB4Mj0iNTQiIHkyPSI0MCIgc3Ryb2tlPSIjMmQzYTRmIiBzdHJva2Utd2lkdGg9IjAuOCIvPgogIDwhLS0gVEVYVE8gU1RFTExBRSDigJQgQSBlbiBnb2xkIC0tPgogIDx0ZXh0IGZvbnQtZmFtaWx5PSJBcmlhbCxIZWx2ZXRpY2Esc2Fucy1zZXJpZiIgZm9udC1zaXplPSIxNSIgZm9udC13ZWlnaHQ9IjUwMCIgbGV0dGVyLXNwYWNpbmc9IjIiPgogICAgPHRzcGFuIHg9IjYyIiB5PSIzMCIgZmlsbD0iI2U3ZTllYSI+U1RFTEw8L3RzcGFuPjx0c3BhbiBmaWxsPSIjQzlBODRDIj5BPC90c3Bhbj48dHNwYW4gZmlsbD0iI2U3ZTllYSI+RTwvdHNwYW4+CiAgPC90ZXh0Pgo8L3N2Zz4="
            style="width:200px; height:48px; display:block; margin-bottom:8px;"
            alt="Stellae">''',
        unsafe_allow_html=True,
    )
    st.divider()

    # =========================================================================
    # SELECTOR GLOBAL DE PROYECTO -- único en toda la app
    # =========================================================================
    st.markdown("**Project**")

    projects_response = (
        supabase.table("projects").select("id, name").order("created_at", desc=True).execute()
    )

    if projects_response.data:
        project_options_sidebar = {p["name"]: p["id"] for p in projects_response.data}
        project_names_sidebar = list(project_options_sidebar.keys())

        # Determinar el index del proyecto activo para pre-seleccionarlo
        _all_options = ["➕ New project..."] + project_names_sidebar
        _active_name = st.session_state.get("project_name", "")
        _sidebar_index = (
            _all_options.index(_active_name)
            if _active_name and _active_name in _all_options
            else 0
        )
        selected_sidebar = st.selectbox(
            "Select project",
            options=_all_options,
            index=_sidebar_index,
            label_visibility="collapsed",
        )

        if selected_sidebar == "➕ New project...":
            project_name = st.text_input("Project name", placeholder="Ex: LNG Plant Phase 2",
                                         key="new_project_name_input")
            project_description = st.text_area("Description (optional)", height=80,
                                               key="new_project_desc_input")
            project_id = None

            # Botón Create Project
            if project_name.strip():
                if st.button("✅ Create Project", type="primary", use_container_width=True,
                             key="btn_create_project"):
                    try:
                        new_pid = get_or_create_project(
                            supabase,
                            project_name.strip(),
                            project_description.strip()
                        )
                        st.session_state.project_name = project_name.strip()
                        st.session_state.project_description = project_description.strip()
                        st.session_state.project_id = new_pid
                        st.success(f"✅ Project '{project_name.strip()}' created!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Could not create project: {e}")
            else:
                st.button("✅ Create Project", type="primary", use_container_width=True,
                          disabled=True, key="btn_create_project_disabled",
                          help="Enter a project name first")
        else:
            project_name = selected_sidebar
            project_id = project_options_sidebar[selected_sidebar]
            project_description = ""
            st.caption(f"ID: {project_id[:8]}...")
            # Badge proyecto activo en sidebar
            st.markdown(
                f'''<div style="
                    background: linear-gradient(135deg, #1E3A5F 0%, #2d5080 100%);
                    border-left: 3px solid #C9A84C;
                    border-radius: 0 6px 6px 0;
                    padding: 6px 10px;
                    margin-top: 8px;
                ">
                    <div style="color:#C9A84C; font-size:9px; letter-spacing:1.5px; text-transform:uppercase; margin-bottom:2px;">Active Project</div>
                    <div style="color:#e7e9ea; font-size:12px; font-weight:600; line-height:1.3;">{project_name}</div>
                </div>''',
                unsafe_allow_html=True
            )
    else:
        st.info("No projects yet. Create your first one.")
        project_name = st.text_input("Project name", placeholder="Ex: LNG Plant Phase 2",
                                     key="first_project_name_input")
        project_description = st.text_area("Description (optional)", height=80,
                                           key="first_project_desc_input")
        project_id = None

        if project_name.strip():
            if st.button("✅ Create Project", type="primary", use_container_width=True,
                         key="btn_create_first_project"):
                try:
                    new_pid = get_or_create_project(
                        supabase,
                        project_name.strip(),
                        project_description.strip()
                    )
                    # Seleccionar el proyecto recién creado automáticamente
                    st.session_state.project_name = project_name.strip()
                    st.session_state.project_description = project_description.strip()
                    st.session_state.project_id = new_pid
                    # Limpiar campos del formulario
                    if "new_project_name_input" in st.session_state:
                        del st.session_state["new_project_name_input"]
                    if "first_project_name_input" in st.session_state:
                        del st.session_state["first_project_name_input"]
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Could not create project: {e}")
        else:
            st.button("✅ Create Project", type="primary", use_container_width=True,
                      disabled=True, key="btn_create_first_project_disabled",
                      help="Enter a project name first")

    st.session_state.project_name = project_name or ""
    st.session_state.project_description = project_description or ""
    st.session_state.project_id = project_id

    st.divider()

    # =========================================================================
    # NAVEGACIÓN -- debajo del selector de proyecto
    # =========================================================================
    current_page = st.radio(
        "Navigation",
        options=["📊 Dashboard", "🔍 Analysis", "🏛️ Governance", "📋 Audit Trail"],
        label_visibility="collapsed",
    )

    render_governance_sidebar_summary()

# =============================================================================
# ROUTER -- Renderiza la página seleccionada en el sidebar
# =============================================================================

if current_page == "📊 Dashboard":
    render_dashboard_page(supabase)
elif current_page == "🔍 Analysis":
    render_analysis_page(supabase)
elif current_page == "🏛️ Governance":
    render_governance_page(supabase)
elif current_page == "📋 Audit Trail":
    render_audit_trail_page(supabase)

