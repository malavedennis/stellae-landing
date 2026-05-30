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
import fitz  # PyMuPDF
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
        "decisions_tab": "Orphan Decisions",
        "changes_tab": "Blind Changes",
        "risks_tab": "Hidden Risks",
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
        "decisions_tab": "Decisiones Huérfanas",
        "changes_tab": "Cambios Ciegos",
        "risks_tab": "Riesgos Ocultos",
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
        "decisions_tab": "Decisões Órfãs",
        "changes_tab": "Mudanças Cegas",
        "risks_tab": "Riscos Ocultos",
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
        "decisions_tab": "Décisions Orphelines",
        "changes_tab": "Changements Aveugles",
        "risks_tab": "Risques Cachés",
    },
}

BASE_SYSTEM_PROMPT = """Actúa como un Auditor Principal de Riesgos y Gobernanza especializado en megaproyectos de infraestructura, energía y capital (con el rigor analítico necesario para evitar fallas catastróficas como las del Aeropuerto de Berlín-Brandenburgo o Crossrail). Tu objetivo es escanear los documentos provistos por el usuario (minutas, reportes, contratos, correspondencia) y extraer de forma cruda, objetiva y sin adornos corporativos únicamente tres tipos de hallazgos latentes. Es crucial que asumas un tono de escepticismo profesional: busca lo que las partes intentan omitir, suavizar o delegar de manera informal.

MICRO-DISPARADORES CRITICOS A DETECTAR (basado en investigacion peer-reviewed de proyectos EPC):
Los siguientes patrones tienen alta correlacion empirica con retrasos de cronograma y sobrecostos en proyectos EPC. Priorizalos en tu analisis:

INGENIERIA Y DISENO:
- Entregables de ingenieria sometidos fuera de plazo o pendientes de aprobacion del cliente sin fecha limite clara
- RFIs (Requests for Information) sin respuesta del cliente por mas de 72 horas
- Cambios en datos de diseno base (rely-upon data) despues del FEED sin analisis de impacto formal
- Long Lead Items (LLI) con TBE (Technical Bid Evaluation) pendiente de aprobacion — bloquea procurement
- P&IDs sin emitir en estado IFC (Issued for Construction) retrasando ingenieria de detalle
- Estudios de HAZOP, HAZID o integridad no completados en FEED y trasladados al EPC

PROCURA Y CONTRATOS:
- Purchase Orders de LLIs colocados tarde respecto al cronograma — define la ruta critica
- Vendor data no entregada a tiempo — fabricacion no puede iniciar sin datos aprobados del proveedor
- Clausulas de penalizacion por demora (liquidated damages) no activadas cuando corresponde
- Facturas o valuaciones pendientes de aprobacion del cliente generando flujo de caja negativo
- Subcontratistas movilizados tarde o con recursos insuficientes vs. plan aprobado

CONSTRUCCION Y CAMPO:
- PTW (Permit to Work) con emision retrasada por parte del cliente en sitios brownfield
- Cambios de alcance implementados sin proceso formal de MOC (Management of Change)
- SIMOPS (Simultaneous Operations) no planificados con anticipacion suficiente
- Interferencias fisicas o de cronograma entre contratistas sin registro de interfaces actualizado
- Recursos de construccion (mano de obra, equipos, maquinaria) insuficientes vs. plan aprobado
- NOC (No Objection Certificate) o permisos de autoridades sin gestionar o con plazo vencido

CONTROL Y MONITOREO:
- Atraso en cronograma detectado sin accion correctiva documentada ni responsable asignado
- Ruta critica impactada sin notificacion formal al cliente ni plan de recuperacion
- Patron de incidentes HSE recurrentes en la misma area o actividad sin escalacion formal
- Lecciones aprendidas de proyectos anteriores no aplicadas — errores repetidos

CIERRE:
- Punch list items pendientes sin fecha de cierre comprometida
- As-built drawings incompletos al momento del cierre mecanico
- Variaciones o reclamaciones contractuales sin resolver al final del proyecto

LOGICA DE CRUCE — REGLAS DE ANALISIS AVANZADO (aplicar siempre):

REGLA 1 — LLI en Ruta Critica:
Cuando detectes retrasos en TBE (Technical Bid Evaluation), Vendor Design, o entrega de Long Lead Items (LLI), evalua si el documento menciona que ese equipo o material esta en la ruta critica (critical path). Si esta en ruta critica: clasifica como RIESGO CRITICO con impacto directo en cronograma — no como riesgo generico. Estima el impacto en semanas si hay datos disponibles.

REGLA 2 — SLA de Aprobacion de Documentos:
Cuando detectes documentos en estado "Pending Client Approval", "Under Client Review", o similares, evalua si el texto menciona cuanto tiempo llevan sin respuesta. Si el plazo supera 72 horas para documentos criticos o 2 semanas para cualquier entregable: activa alerta de incumplimiento de SLA contractual. El paper SPE-203431 establece que documentos sin respuesta dentro del ciclo contractual deben considerarse aprobados por defecto — registra esta exposicion del cliente.

REGLA 3 — PTW como indicador de productividad perdida:
Cuando detectes menciones de retrasos en Permit to Work (PTW), "work front availability", o permisos de trabajo pendientes: evalua si es un patron recurrente (mas de una mencion en el mismo documento o referencia a semanas anteriores). Si es recurrente: calcula o estima el impacto acumulado en productividad de campo (horas-hombre perdidas por turno si hay datos disponibles).

REGLA 4 — FEED deficiencies como riesgo de alcance:
Cuando detectes menciones de "FEED shortages", "FEED verification", "rely upon data", discrepancias en datos base de diseno, o estudios que debieron completarse en FEED pero se trasladaron al EPC: activa bandera de Riesgo de Cambio de Alcance Contractual. Este patron tiene alta correlacion empirica con disputes y change orders post-adjudicacion.

REGLA 5 — Plan de 90 dias como hito auditable:
Si el proyecto esta en etapa de inicio o las minutas corresponden a las primeras semanas de ejecucion EPC: verifica si el documento menciona la existencia de un plan de 90 dias (90-day frontend plan). Si no hay evidencia de este plan: registralo como decision huerfana — ausencia de planificacion frontal documentada es un predictor de problemas en etapas posteriores.

DICCIONARIO DE MICRO-DISPARADORES POR INDUSTRIA:
Cuando el contexto del proyecto indique una industria especifica, aplica los micro-disparadores adicionales correspondientes:

SI la industria es OIL & GAS / UPSTREAM / DOWNSTREAM / PETROCHEMICAL / LNG:
- HAZOP o HAZID no completado o con cambios de diseno posteriores no re-evaluados
- SIS (Safety Instrumented System) o SIL rating modificado sin revision formal
- Cambio en perfiles de produccion o composicion de fluidos sin actualizar la ingenieria base
- Pipeline o flowline corridor study no concluido antes del EPC
- Brownfield tie-ins no verificados fisicamente en campo (as-built vs. design)
- Cambio en diametro, presion de diseno, o temperatura de operacion sin actualizar P&IDs

SI la industria es MINERIA / MINING / MINERALES:
- Estudios geotecnicos no completados o con supuestos de diseno no validados en campo
- Cambio en angulo de talud o metodo de extraccion sin analisis de estabilidad formal
- Environmental Assessment (EA) o permiso ambiental sin gestion activa o con plazo vencido
- Plan de manejo de relaves (tailings) modificado sin aprobacion del ente regulatorio
- Community agreement o consulta indigena no documentada antes de actividades de campo
- Cambio en ley de mineral o en metodo de procesamiento sin impacto en el diseno evaluado
- Freeze de precios de commodities como trigger de scope review — proyectos pausados sin decision formal

SI la industria es INFRAESTRUCTURA / INFRASTRUCTURE / CONSTRUCCION / VIALIDAD / PUENTES:
- Right-of-way (derecho de via) no asegurado o con oposicion legal sin resolver
- Interferencia con servicios publicos (utilities) no notificada o no coordinada formalmente
- Environmental Impact Statement (EIS) con cambios de diseno no incorporados
- Aprobacion politica o regulatoria pendiente que bloquea frentes de trabajo
- Conflictos de coordinacion entre multiples contratistas sin registro de interfaces actualizado
- Relocalizacion de comunidades o afectaciones sociales sin plan formal documentado

SI la industria es ENERGIA / POWER / TRANSMISION / RENOVABLES:
- Modificacion de setpoints SIS sin revision de seguridad funcional (IEC 61511)
- Cambio de equipo principal (transformador, turbina, generador) sin actualizar estudios de flujo de potencia
- Grid connection agreement con el operador de red sin firmar o con condiciones pendientes
- Permiso de operacion comercial (COD) con hitos tecnicos sin completar
- Cambio en layout de planta solar o eolica sin re-evaluacion de sombras o viento

SI la industria es CONSTRUCCION GENERAL / EPC / EPCM (sin industria especifica):
Aplicar todos los micro-disparadores generales de las secciones anteriores de este prompt.

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


def _extract_text_ocr_fallback(pdf_bytes: bytes, ocr_pages: list = None) -> str:
    """
    OCR fallback para páginas escaneadas (sin capa de texto digital).
    Usa pdf2image + pytesseract — mismo motor que antes.
    ocr_pages: lista de índices de página (0-based). None = todas las páginas.
    """
    try:
        _kwargs = dict(dpi=200)
        if ocr_pages is not None and len(ocr_pages) > 0:
            _kwargs["first_page"] = min(ocr_pages) + 1
            _kwargs["last_page"]  = max(ocr_pages) + 1
        if POPPLER_PATH:
            _kwargs["poppler_path"] = POPPLER_PATH
        images = convert_from_bytes(pdf_bytes, **_kwargs)
        texts = []
        for img in images:
            ocr_text = extract_text_with_ocr(img)
            if ocr_text:
                texts.append(f"[OCR] {ocr_text}")
        return "\n\n".join(texts)
    except Exception as e:
        return f"[OCR unavailable: {e}]"


def extract_text_from_pdf(uploaded_file) -> str:
    """
    Extrae texto de un PDF usando PyMuPDF (fitz).
    Más fiel que PyPDF2 para layouts complejos, tablas y documentos EPC.

    Estrategia por página:
    - Si la página tiene texto digital (>30 chars) → PyMuPDF directo.
    - Si la página está escaneada (sin texto o <30 chars) → OCR fallback.
    Esto preserva la lógica híbrida que ya existía con PyPDF2.
    """
    pdf_bytes = uploaded_file.getvalue()

    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")

        pages_text    = []   # (page_index, text)
        ocr_needed    = []   # índices de páginas que necesitan OCR

        for i, page in enumerate(doc):
            page_text = page.get_text().strip()
            if len(page_text) > 30:
                pages_text.append((i, page_text))
            else:
                ocr_needed.append(i)

        doc.close()

        # Páginas escaneadas — aplicar OCR
        if ocr_needed:
            ocr_result = _extract_text_ocr_fallback(pdf_bytes, ocr_pages=ocr_needed)
            if ocr_result:
                # Insertar OCR al final con marcador de página
                pages_text.append((max(ocr_needed), ocr_result))

        # Ordenar por número de página y unir
        pages_text.sort(key=lambda x: x[0])
        return "\n".join(text for _, text in pages_text)

    except Exception as e:
        # Último recurso: OCR completo del PDF
        try:
            return _extract_text_ocr_fallback(pdf_bytes)
        except Exception:
            return f"[PDF extraction error: {e}]"


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


def extract_all_text(uploaded_files) -> tuple:
    """
    Extrae texto de todos los archivos subidos.
    Soporta: PDF (digital + escaneado con OCR), DOCX, TXT, JPG, PNG.

    DOBLE VÍA (implementado Mayo 2026):
    Retorna dos versiones del texto para distintos usos:
    - raw_text:        Texto original con acentos y mayúsculas intactas.
                       Se envía a Claude — preserva semántica para mejor comprensión del LLM.
    - normalized_text: Texto limpio (minúsculas, sin acentos).
                       Se usa SOLO para búsqueda de micro-triggers por keyword.

    Returns:
        tuple(raw_text: str, normalized_text: str)
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

    raw_text = "\n\n".join(sections)
    normalized_text = normalize_text(raw_text)
    return raw_text, normalized_text


# =============================================================================
# CHUNKING SEMÁNTICO
# =============================================================================

def chunk_text_smart(
    text: str,
    max_chars: int = MAX_CHARS,
    overlap_chars: int = 500,
) -> list:
    """
    Divide texto extenso en chunks respetando límites naturales de párrafo.
    Reemplaza el truncado rígido por caracteres (raw_text[:MAX_CHARS]).

    Mejoras vs truncado rígido:
    - Corta en párrafos naturales (\n\n), no a mitad de oración
    - Agrega overlap entre chunks — el inicio de cada chunk incluye el
      final del anterior para no perder contexto entre segmentos
    - Garantiza que todos los documentos se analizan, no solo los primeros
      MAX_CHARS caracteres

    Args:
        text:         Texto completo a dividir (raw o normalized — ambos sirven).
        max_chars:    Tamaño máximo de cada chunk en caracteres.
        overlap_chars: Caracteres de solapamiento entre chunks para preservar contexto.

    Returns:
        list[str]: Lista de chunks. Si el texto cabe en un solo chunk, lista de 1 elemento.
    """
    if len(text) <= max_chars:
        return [text]

    # Dividir en párrafos naturales (doble salto de línea)
    paragraphs = [p.strip() for p in text.split("\n\n") if p.strip()]

    chunks   = []
    current  = []
    current_len = 0

    for para in paragraphs:
        para_len = len(para)

        if current_len + para_len > max_chars and current:
            # Cerrar chunk actual
            chunk_text = "\n\n".join(current)
            chunks.append(chunk_text)

            # Overlap: conservar el final del chunk anterior como contexto
            # para que el inicio del próximo no pierda el hilo
            if overlap_chars > 0:
                overlap_text = chunk_text[-overlap_chars:]
                current     = [f"[...contexto anterior...]\n{overlap_text}", para]
                current_len = len(overlap_text) + para_len + 30
            else:
                current     = [para]
                current_len = para_len
        else:
            current.append(para)
            current_len += para_len

    # Último chunk
    if current:
        chunks.append("\n\n".join(current))

    return chunks


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
{document_text[:40000]}

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
        # Limpiar caracteres que se renderizan como cuadros corruptos
        content = content.replace("■", "●").replace("▨", "●").replace("\u25a0", "●").replace("\u25a8", "●")
        if finding.get("governance_violation"):
            violated_rule = finding.get("violated_rule") or "Unknown rule"
            # Traducir violated_rule al idioma UI — siempre, sin depender de detect_content_language
            _ui_lang = st.session_state.get("output_language", "en")
            _lang_names = {"en": "English", "es": "Spanish", "pt": "Portuguese", "fr": "French"}
            _target_lang_name = _lang_names.get(_ui_lang, "English")
            try:
                _translated_rule = anthropic.Anthropic(
                    api_key=os.getenv("ANTHROPIC_API_KEY")
                ).messages.create(
                    model="claude-haiku-4-5-20251001",
                    max_tokens=100,
                    messages=[{"role": "user", "content":
                        f"Translate this governance rule name to {_target_lang_name}. Return ONLY the translation, no explanation: {violated_rule}"}]
                ).content[0].text.strip()
                violated_rule = _translated_rule
            except Exception:
                pass  # usar original si falla
            # Header GOVERNANCE VIOLATION traducido por idioma
            _viol_headers = {
                "en": "GOVERNANCE VIOLATION — Escalation required",
                "es": "VIOLACIÓN DE GOBERNANZA — Escalación requerida",
                "pt": "VIOLAÇÃO DE GOVERNANÇA — Escalada necessária",
                "fr": "VIOLATION DE GOUVERNANCE — Escalade requise",
            }
            _viol_header = _viol_headers.get(_ui_lang, _viol_headers["en"])
            st.markdown(
                f'''<div style="background:rgba(201,50,50,0.12);border:1px solid rgba(201,50,50,0.4);
                border-left:4px solid #c93232;border-radius:4px;padding:10px 16px;margin-bottom:8px;">
                <span style="color:#ff6b6b;font-weight:700;font-size:13px;">
                {_viol_header}</span><br>
                <span style="color:#ffaaaa;font-size:12px;">{violated_rule}</span>
                </div>''',
                unsafe_allow_html=True
            )
            st.markdown(content)
            st.divider()
        elif category == "decision":
            st.markdown(
                f'''<div style="color:#e7e9ea;border-left:3px solid #C9A84C;
                padding-left:12px;margin-bottom:4px;">
                <span style="color:#C9A84C;font-weight:600;font-size:12px;">DECISION</span><br>
                {content}</div>''',
                unsafe_allow_html=True
            )
            st.divider()
        elif category == "change":
            st.markdown(
                f'''<div style="color:#e7e9ea;border-left:3px solid #60a5fa;
                padding-left:12px;margin-bottom:4px;">
                <span style="color:#60a5fa;font-weight:600;font-size:12px;">CHANGE DETECTED</span><br>
                {content}</div>''',
                unsafe_allow_html=True
            )
            st.divider()
        elif category == "risk":
            st.markdown(
                f'''<div style="color:#e7e9ea;border-left:3px solid #f87171;
                padding-left:12px;margin-bottom:4px;">
                <span style="color:#f87171;font-weight:600;font-size:12px;">RISK IDENTIFIED</span><br>
                {content}</div>''',
                unsafe_allow_html=True
            )
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
    text = text.replace("■", "●").replace("▨", "●").replace("\u25a0", "●").replace("\u25a8", "●")
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
    """Traduce findings al idioma del PDF. Evalua cada finding individualmente."""
    if not findings:
        return findings

    # Separar findings con y sin traducción — evaluar CADA finding por separado
    needs_translation = []
    needs_translation_idx = []
    result = list(findings)  # copia para modificar

    for i, finding in enumerate(findings):
        content = finding.get("content", "")
        content_lang = detect_content_language(content)

        # Si el contenido ya está en el idioma target — no necesita traducción
        if content_lang == target_lang:
            continue

        # Verificar cache válido para este finding específico
        translations = finding.get("content_translations") or {}
        cached = translations.get(target_lang)
        cache_lang = detect_content_language(cached) if cached else "unknown"
        cache_valid = (
            cached
            and isinstance(cached, str)
            and len(cached.strip()) > 20
            and cache_lang == target_lang
        )

        if cache_valid:
            result[i] = dict(finding)
            result[i]["content"] = cached
        else:
            needs_translation.append(content)
            needs_translation_idx.append(i)

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
            rule_clean = f.get("violated_rule", "N/A")
            # Traducir violated_rule al idioma del PDF
            _pdf_lang_names = {"en": "English", "es": "Spanish", "pt": "Portuguese", "fr": "French"}
            _pdf_target = _pdf_lang_names.get(language, "English")
            try:
                _client_pdf = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
                rule_clean = _client_pdf.messages.create(
                    model="claude-haiku-4-5-20251001", max_tokens=100,
                    messages=[{"role": "user", "content":
                        f"Translate this governance rule name to {_pdf_target}. Return ONLY the translation: {rule_clean}"}]
                ).content[0].text.strip()
            except Exception:
                pass
            rule_clean = clean_for_pdf(rule_clean)
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
        if not text or len(text.strip()) < 20:
            return text
        # Detectar textos vacíos/sin hallazgos en cualquier idioma (no solo EN/ES)
        _no_findings_markers = [
            "No se detectaron", "No findings", "No anomalies",
            "Aucune anomalie", "Nenhuma anomalia"
        ]
        if any(marker in text for marker in _no_findings_markers):
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
            _sb = st.session_state.get("supabase_client") or supabase
            findings_list = apply_governance_rules(project_id, findings_list, _sb)
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

    # Labels de tabs con iconos según el idioma activo
    _cur_lang = st.session_state.get("output_language", "en")
    _L = STRUCTURE_LABELS.get(_cur_lang, STRUCTURE_LABELS["en"])
    tab_decisiones, tab_cambios, tab_riesgos = st.tabs([
        f"🛡️ {_L['decisions_tab']}",
        f"🔄 {_L['changes_tab']}",
        f"⚠️ {_L['risks_tab']}",
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


# =============================================================================
# STELLAE PREDICTIVE RISK ENGINE — SEM Phase 1 (Deterministic Academic Priors)
# Based on peer-reviewed research. No training data required.
# Sources:
#   - Shen, Tang, Wang, Duffield et al. (2021) — J. Construction Engineering
#     and Management. n=85 megaprojects, 4 continents. β=0.88, p<0.01
#   - Flyvbjerg (2007) — 258 infrastructure projects, 20 nations, 70 years
#   - SPE-203215-MS — Cost commitment by phase in EPC projects
#   - McKinsey Global Institute (2016) — 2,500+ construction projects
# =============================================================================

ACADEMIC_PRIORS = {
    # Shen et al. 2021 — governance formal → interface management performance
    "governance_to_interface": 0.88,        # path coefficient β
    "governance_explains_variance": 0.89,   # R² = 89%

    # Flyvbjerg 2007 — base probability of cost overrun
    "base_overrun_probability": 0.86,       # 9 of 10 megaprojects

    # SPE-203215-MS — cost of a late decision multiplied by project phase
    # A decision made in Construction costs 25x more than in FEED
    # Keys use lowercase_underscore to match project_phase sanitizer output:
    #   project_phase.lower().replace(" ", "_").replace("-", "_")
    "decision_cost_multiplier": {
        "concept":              1.0,
        "feed":                 3.5,
        "detailed_engineering": 8.0,
        "construction":         25.0,
        "commissioning":        60.0,
        # Common aliases from Supabase project_stage field:
        "pre_feed":             2.0,
        "basic_engineering":    5.0,
        "procurement":          15.0,
        "pre_commissioning":    45.0,
    },

    # McKinsey Global Institute 2016 — industry benchmarks
    "projects_over_budget_pct": 0.57,
    "projects_delayed_pct":     0.77,
}

# =============================================================================
# FLYVBJERG REFERENCE CLASS FORECASTING — P80 overrun by project class
# Source: Flyvbjerg (2007), n=258, 70 years. P80 = 80th percentile upper bound.
# Used by calculate_pert_coi() to set the pessimistic scenario of the PERT range.
# Keys match industry/project_type fields from Project Context in Supabase.
# =============================================================================
FLYVBJERG_P80 = {
    # Oil & Gas / Energy
    "lng_petrochemical":    0.80,   # LNG, gas plants, refineries ~35-45% avg, P80 ~80%
    "oil_gas":              0.80,   # General O&G upstream/downstream
    "lng":                  0.80,
    "petrochemical":        0.80,
    "refinery":             0.75,
    "offshore":             0.85,   # Offshore projects have higher variance
    # Mining
    "mining":               0.90,   # Mining expansion ~45% avg, P80 ~90%
    "mining_expansion":     0.90,
    "copper":               0.90,
    "lithium":              0.95,
    # Civil / Infrastructure
    "tunnel":               1.20,   # Tunnels ~66% avg, P80 ~120%
    "tunnel_underground":   1.20,
    "bridge":               0.65,   # Bridges ~34% avg, P80 ~65%
    "bridge_viaduct":       0.65,
    "rail":                 0.85,   # Rail/metro ~45% avg
    "metro":                0.85,
    "rail_metro":           0.85,
    "road":                 0.55,   # Roads ~20% avg, lower variance
    # General
    "infrastructure":       0.65,   # General infrastructure ~34% avg
    "construction":         0.65,
    "epc":                  0.75,   # Generic EPC
    "default":              0.75,   # Conservative default if class unknown
}


def calculate_pert_coi(
    burn_rate_usd_day: float,
    days_open: int,
    phase_multiplier: float,
    interference_multiplier: float,
    project_class: str = "default",
) -> dict:
    """
    Beta-PERT Cost of Inaction — probabilistic range instead of a single point.

    Replaces the deterministic COI with a three-point PERT estimate:
    - Optimistic: minimum interference, no Flyvbjerg overrun
    - Mode: current interference multiplier (same as Stage 1 deterministic)
    - Pessimistic: Flyvbjerg P80 upper bound for the project class

    PERT Mean = (Optimistic + 4×Mode + Pessimistic) / 6
    This is the standard PMI Beta-PERT formula (MacCrimmon & Ryavec, RAND 1964).

    Args:
        burn_rate_usd_day:      Daily project burn rate in USD.
        days_open:              Average days findings have been open.
        phase_multiplier:       SPE-203215 phase cost multiplier.
        interference_multiplier: Category-based interference multiplier.
        project_class:          Project type key from FLYVBJERG_P80 table.
                                Mapped from project_context industry/project_type fields.

    Returns:
        dict with coi_min, coi_mode, coi_max, coi_mean, coi_range_str,
        flyvbjerg_p80_pct, project_class_used.

    Stage 1 compatibility: coi_mode == calculate_governance_risk()["inaction_cost_usd"]
    The mode is identical to the Stage 1 deterministic value — no regression.
    """
    base = burn_rate_usd_day * days_open * phase_multiplier

    coi_optimistic  = base * 1.0                    # Best case: no interference
    coi_mode        = base * interference_multiplier  # Current state (= Stage 1 value)

    # Pessimistic: Flyvbjerg P80 applied as additional multiplier
    # P80 of 0.80 means "80% chance overrun stays below 80% above mode"
    p80 = FLYVBJERG_P80.get(project_class.lower().replace(" ", "_").replace("-", "_"),
                             FLYVBJERG_P80["default"])
    coi_pessimistic = coi_mode * (1 + p80)

    # Standard PERT mean formula
    coi_mean = (coi_optimistic + 4 * coi_mode + coi_pessimistic) / 6

    def _fmt(v: float) -> str:
        if v >= 1_000_000:
            return f"${v/1_000_000:.1f}M"
        elif v >= 1_000:
            return f"${v/1_000:.0f}K"
        return f"${v:.0f}"

    return {
        "coi_min":            round(coi_optimistic, 0),
        "coi_mode":           round(coi_mode, 0),
        "coi_max":            round(coi_pessimistic, 0),
        "coi_mean":           round(coi_mean, 0),
        "coi_range_str":      f"{_fmt(coi_optimistic)} – {_fmt(coi_pessimistic)}",
        "coi_central_str":    _fmt(coi_mode),
        "coi_mean_str":       _fmt(coi_mean),
        "flyvbjerg_p80_pct":  round(p80 * 100, 0),
        "project_class_used": project_class,
    }


def _map_project_class(industry: str = None, project_type: str = None) -> str:
    """
    Maps the industry / project_type fields from Project Context to a
    FLYVBJERG_P80 key. Tries project_type first, then industry, then default.
    """
    candidates = []
    if project_type:
        candidates.append(project_type.lower().replace(" ", "_").replace("-", "_"))
    if industry:
        candidates.append(industry.lower().replace(" ", "_").replace("-", "_"))

    for c in candidates:
        if c in FLYVBJERG_P80:
            return c
        # Partial match — e.g. "oil & gas" → "oil_gas"
        for key in FLYVBJERG_P80:
            if key in c or c in key:
                return key

    return "default"


def calculate_interference_multiplier(open_findings: list) -> float:
    """
    Interference Multiplier — diferencia el COI según el tipo de finding.
    
    COI = burn_rate × days_open × phase_multiplier × interference_multiplier
    
    Methodology (METHODOLOGY.md):
    - 1.0x: decisión huérfana o riesgo sin impacto físico
    - 1.5x: cambio sin CR — riesgo de rework potencial
    - 2.0x: riesgo HSE o patrón sistémico — paralización potencial
    - 2.5x: cambio confirmado con rework en campo (demolición, re-procura)
    
    En Stage 2 se calibrará con datos reales de cierres de proyectos.
    """
    if not open_findings:
        return 1.0
    
    # Mapeo por categoría de finding
    CATEGORY_MULTIPLIERS = {
        "decision": 1.0,   # Decisión huérfana — impacto documental
        "change":   1.5,   # Cambio ciego — riesgo de rework físico
        "risk":     2.0,   # Riesgo oculto — paralización potencial
    }
    
    # Tomar el multiplicador más alto de los findings activos
    # (el peor escenario activo domina el COI)
    max_multiplier = 1.0
    for f in open_findings:
        cat = f.get("category", "decision")
        m = CATEGORY_MULTIPLIERS.get(cat, 1.0)
        # Si hay violation activa, escalar 0.5x adicional
        if f.get("governance_violation") and f.get("status") == "open":
            m = min(2.5, m + 0.5)
        max_multiplier = max(max_multiplier, m)
    
    return max_multiplier


def calculate_governance_risk(
    violations: int,
    project_phase: str,
    days_open: int,
    burn_rate_usd_day: float,
    open_findings: list = None,
) -> dict:
    """
    Stellae Predictive Risk Engine — Phase 1 (Deterministic).

    Propagates governance violations through academic coefficients to produce
    quantified financial risk estimates. No training data required — uses
    peer-reviewed priors as fixed weights.

    COI Formula (METHODOLOGY.md):
        COI = burn_rate × days_open × phase_multiplier × interference_multiplier

    Args:
        violations: number of active governance violations
        project_phase: one of concept/feed/detailed_engineering/construction/commissioning
        days_open: average days the findings have been open without resolution
        burn_rate_usd_day: project daily expenditure rate in USD
        open_findings: list of finding dicts — used to calculate interference multiplier

    Returns:
        dict with governance_health_pct, interface_risk_pct,
              inaction_cost_usd, overrun_probability_pct, citation,
              interference_multiplier
    """
    # Governance health — protected floor at 5% (never reaches absolute zero)
    VIOLATION_WEIGHTS = {
        "lli_critical_path":    0.25,
        "hse_pattern":          0.20,
        "change_no_cr":         0.18,
        "sla_document_overdue": 0.15,
        "orphan_decision":      0.12,
        "feed_deficiency":      0.10,
        "default":              0.12,
    }
    DECAY_CONSTANT = VIOLATION_WEIGHTS["default"]
    governance_health = max(0.05, 1.0 - (violations * DECAY_CONSTANT))

    # Interface risk — propagated via Shen et al. β=0.88
    interface_risk = 1.0 - (governance_health * ACADEMIC_PRIORS["governance_to_interface"])

    # Phase multiplier — SPE-203215
    phase_multiplier = ACADEMIC_PRIORS["decision_cost_multiplier"].get(
        project_phase.lower().replace(" ", "_").replace("-", "_"),
        1.0
    )

    # Interference multiplier — based on finding categories (METHODOLOGY.md)
    interference_multiplier = calculate_interference_multiplier(open_findings or [])

    # Cost of Inaction — fórmula completa con interference multiplier
    # COI = burn_rate × days_open × phase_multiplier × interference_multiplier
    inaction_cost = burn_rate_usd_day * days_open * phase_multiplier * interference_multiplier

    # Overrun probability — bounded at 95% ceiling
    overrun_prob = min(
        0.95,
        ACADEMIC_PRIORS["base_overrun_probability"] * interface_risk
    )

    return {
        "governance_health_pct":    round(governance_health * 100, 1),
        "interface_risk_pct":       round(interface_risk * 100, 1),
        "inaction_cost_usd":        round(inaction_cost, 0),
        "overrun_probability_pct":  round(overrun_prob * 100, 1),
        "phase_multiplier":         phase_multiplier,
        "interference_multiplier":  interference_multiplier,
        "citation": (
            "Shen et al. 2021 (β=0.88, n=85) · "
            "Flyvbjerg 2007 (n=258, 70yr) · "
            f"SPE-203215 ({phase_multiplier}x {project_phase}) · "
            f"Interference {interference_multiplier}x"
        ),
        # Beta-PERT range — populated by render_predictive_risk_panel()
        # after project_class is resolved from Project Context.
        "coi_pert": None,
    }


def render_predictive_risk_panel(
    all_findings: list,
    project_phase: str,
    burn_rate_usd_day: float = 50000.0,
    project_class: str = "default",
) -> None:
    """
    Renders the Predictive Risk Panel in the Dashboard.
    Uses calculate_governance_risk() with academic priors (Stage 1 deterministic).
    Beta-PERT range via calculate_pert_coi() using Flyvbjerg P80 by project class.

    Args:
        all_findings:       All findings for the current project.
        project_phase:      Current project phase (maps to SPE-203215 multiplier).
        burn_rate_usd_day:  Daily burn rate in USD. Read from Project Context; fallback $50K.
        project_class:      Flyvbjerg reference class key (from industry/project_type fields).
                            Determines the P80 pessimistic scenario of the PERT COI range.
    """
    if not all_findings:
        return

    # Count active violations and open findings
    active_violations = sum(
        1 for f in all_findings
        if f.get("governance_violation") and f.get("status") == "open"
    )
    open_findings = [f for f in all_findings if f.get("status") == "open"]

    # Average days open (use 1 as minimum to avoid zero cost display)
    if open_findings:
        import datetime
        days_list = []
        for f in open_findings:
            created = f.get("created_at", "")
            if created:
                try:
                    created_dt = datetime.datetime.fromisoformat(
                        created.replace("Z", "+00:00")
                    )
                    delta = (
                        datetime.datetime.now(datetime.timezone.utc) - created_dt
                    ).days
                    days_list.append(max(1, delta))
                except Exception:
                    days_list.append(1)
        avg_days_open = sum(days_list) / len(days_list) if days_list else 1
    else:
        avg_days_open = 1

    if active_violations == 0 and not open_findings:
        return

    risk = calculate_governance_risk(
        violations=active_violations,
        project_phase=project_phase,
        days_open=int(avg_days_open),
        burn_rate_usd_day=burn_rate_usd_day,
        open_findings=open_findings,
    )

    # Beta-PERT COI range — Flyvbjerg Reference Class Forecasting
    pert = calculate_pert_coi(
        burn_rate_usd_day=burn_rate_usd_day,
        days_open=int(avg_days_open),
        phase_multiplier=risk["phase_multiplier"],
        interference_multiplier=risk["interference_multiplier"],
        project_class=project_class,
    )
    risk["coi_pert"] = pert

    # Labels traducidos por idioma activo
    _ui_lang = st.session_state.get("output_language", "en")
    _panel_labels = {
        "en": {
            "title": "Governance Risk Estimate",
            "health": "Governance Health",
            "interface": "Interface Risk",
            "cost": "Cost of Inaction",
            "cost_range_label": "PERT Range",
            "cost_central_label": "central",
            "overrun": "Overrun Est. Risk",
            "disclaimer": (
                "Directional estimates based on academic priors, not statistically certified predictions. "
                "Accuracy improves as Stellae accumulates real project outcomes."
            ),
            "tooltip": (
                "Academic baseline · Shen et al. 2021 (β=0.88 propagation factor, n=85) · "
                "Flyvbjerg 2007 Reference Class Forecasting (n=258, 70yr) · SPE-203215 phase multiplier · "
                "Beta-PERT range uses Flyvbjerg P80 as pessimistic scenario · "
                "Stage 1 deterministic engine — Bayesian calibration activates in Stage 2"
            ),
        },
        "es": {
            "title": "Estimación de Riesgo de Gobernanza",
            "health": "Salud de Gobernanza",
            "interface": "Riesgo de Interfaz",
            "cost": "Costo de la Inacción",
            "cost_range_label": "Rango PERT",
            "cost_central_label": "central",
            "overrun": "Riesgo Est. de Sobrecosto",
            "disclaimer": (
                "Estimaciones directivas basadas en priors académicos, no predicciones estadísticamente certificadas. "
                "La precisión mejora conforme Stellae acumula resultados reales de proyectos."
            ),
            "tooltip": (
                "Baseline académico · Shen et al. 2021 (β=0.88 factor de propagación, n=85) · "
                "Flyvbjerg 2007 Reference Class Forecasting (n=258, 70 años) · Multiplicador de fase SPE-203215 · "
                "Rango Beta-PERT usa P80 de Flyvbjerg como escenario pesimista · "
                "Motor determinista Stage 1 — calibración bayesiana se activa en Stage 2"
            ),
        },
        "pt": {
            "title": "Estimativa de Risco de Governança",
            "health": "Saúde da Governança",
            "interface": "Risco de Interface",
            "cost": "Custo da Inação",
            "cost_range_label": "Faixa PERT",
            "cost_central_label": "central",
            "overrun": "Risco Est. de Sobrecusto",
            "disclaimer": (
                "Estimativas direcionais baseadas em priors acadêmicos, não previsões certificadas estatisticamente. "
                "A precisão melhora conforme a Stellae acumula resultados reais de projetos."
            ),
            "tooltip": (
                "Baseline acadêmico · Shen et al. 2021 (β=0.88 fator de propagação, n=85) · "
                "Flyvbjerg 2007 Reference Class Forecasting (n=258, 70 anos) · SPE-203215 · "
                "Faixa Beta-PERT usa P80 de Flyvbjerg como cenário pessimista · "
                "Motor determinístico Stage 1 — calibração bayesiana ativa no Stage 2"
            ),
        },
        "fr": {
            "title": "Estimation du Risque de Gouvernance",
            "health": "Santé de Gouvernance",
            "interface": "Risque d'Interface",
            "cost": "Coût de l'Inaction",
            "cost_range_label": "Plage PERT",
            "cost_central_label": "central",
            "overrun": "Risque Est. de Dépassement",
            "disclaimer": (
                "Estimations directionnelles basées sur des priors académiques, non des prédictions certifiées. "
                "La précision s'améliore au fur et à mesure que Stellae accumule des résultats réels."
            ),
            "tooltip": (
                "Baseline académique · Shen et al. 2021 (β=0.88 facteur de propagation, n=85) · "
                "Flyvbjerg 2007 Reference Class Forecasting (n=258, 70 ans) · SPE-203215 · "
                "Plage Beta-PERT utilise P80 de Flyvbjerg comme scénario pessimiste · "
                "Moteur déterministe Stage 1 — calibration bayésienne activée en Stage 2"
            ),
        },
    }
    _lbl = _panel_labels.get(_ui_lang, _panel_labels["en"])

    # Título + ℹ️ con st.popover (funciona en Streamlit como click)
    _title_col, _info_col = st.columns([8, 1])
    with _title_col:
        st.markdown(f"#### 🔬 {_lbl['title']}")
    with _info_col:
        with st.popover("ℹ️"):
            st.caption(_lbl['tooltip'])

    col1, col2, col3, col4 = st.columns(4)

    # Estilo común para todas las tarjetas — altura uniforme con flexbox
    _card_base = (
        "text-align:center;background:rgba(255,255,255,0.04);"
        "border:1px solid rgba(255,255,255,0.08);border-radius:8px;padding:20px 8px;"
        "min-height:110px;display:flex;flex-direction:column;justify-content:center;"
    )

    with col1:
        health = risk["governance_health_pct"]
        color = "#4cb87a" if health >= 70 else "#C9A84C" if health >= 40 else "#ff6b6b"
        st.markdown(
            f'''<div style="{_card_base}">
            <div style="font-size:28px;font-weight:800;color:{color};">{health}%</div>
            <div style="font-size:11px;color:#9a9690;margin-top:4px;">{_lbl["health"]}</div>
            </div>''',
            unsafe_allow_html=True
        )

    with col2:
        irisk = risk["interface_risk_pct"]
        color2 = "#4cb87a" if irisk <= 15 else "#C9A84C" if irisk <= 35 else "#ff6b6b"
        st.markdown(
            f'''<div style="{_card_base}">
            <div style="font-size:28px;font-weight:800;color:{color2};">{irisk}%</div>
            <div style="font-size:11px;color:#9a9690;margin-top:4px;">{_lbl["interface"]}</div>
            </div>''',
            unsafe_allow_html=True
        )

    with col3:
        cost = risk["inaction_cost_usd"]
        pert = risk.get("coi_pert")
        if pert:
            # Beta-PERT range — misma altura que las otras tarjetas
            cost_str = pert["coi_range_str"]
            cost_lbl = f"{_lbl.get('cost_range_label','PERT Range')} · {_lbl.get('cost_central_label','central')}: {pert['coi_central_str']}"
            st.markdown(
                f'''<div style="{_card_base}">
                <div style="font-size:22px;font-weight:800;color:#ff6b6b;line-height:1.2;letter-spacing:-0.5px;">{cost_str}</div>
                <div style="font-size:10px;color:#9a9690;margin-top:5px;line-height:1.3;">{cost_lbl}</div>
                </div>''',
                unsafe_allow_html=True
            )
        else:
            # Fallback: Stage 1 deterministic single value
            cost_str = f"${cost/1e6:.1f}M" if cost >= 1e6 else f"${cost/1e3:.0f}K"
            st.markdown(
                f'''<div style="{_card_base}">
                <div style="font-size:28px;font-weight:800;color:#ff6b6b;">{cost_str}</div>
                <div style="font-size:11px;color:#9a9690;margin-top:4px;">{_lbl["cost"]}</div>
                </div>''',
                unsafe_allow_html=True
            )

    with col4:
        prob = risk["overrun_probability_pct"]
        color4 = "#4cb87a" if prob <= 30 else "#C9A84C" if prob <= 60 else "#ff6b6b"
        st.markdown(
            f'''<div style="{_card_base}">
            <div style="font-size:28px;font-weight:800;color:{color4};">{prob}%</div>
            <div style="font-size:11px;color:#9a9690;margin-top:4px;">{_lbl["overrun"]}</div>
            </div>''',
            unsafe_allow_html=True
        )

    st.caption(
        f"⚠️ {_lbl['disclaimer']} "
        f"· Phase: {risk['phase_multiplier']}x · Interference: {risk['interference_multiplier']}x"
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

    # --- Predictive Risk Panel al tope (SEM Phase 1 — Academic Priors) ---
    # Leer burn_rate y phase desde Project Context; fallback a valores por defecto
    _proj_ctx = load_project_context(supabase_client, project_id)
    _proj_stage = _proj_ctx.get("project_stage", "construction") if _proj_ctx else "construction"
    _proj_phase = (_proj_stage or "construction").lower().replace(" ", "_").replace("-", "_")
    _burn_rate_raw = _proj_ctx.get("burn_rate_usd_day", None)
    try:
        _burn_rate = float(_burn_rate_raw) if _burn_rate_raw else 50000.0
    except (ValueError, TypeError):
        _burn_rate = 50000.0
    # Resolver project_class desde el Project Context para Beta-PERT
    _proj_industry = _proj_ctx.get("industry", "default") if _proj_ctx else "default"
    _proj_type     = _proj_ctx.get("project_type", "") if _proj_ctx else ""
    _proj_class    = _map_project_class(
        industry=_proj_industry,
        project_type=_proj_type,
    )
    render_predictive_risk_panel(
        all_findings=all_findings,
        project_phase=_proj_phase,
        burn_rate_usd_day=_burn_rate,
        project_class=_proj_class,
    )

    # Calcular semáforo
    status_label, status_message, status_level = calculate_project_status(all_findings)

    # --- Sección 1: Semáforo prominente + botón de reporte al lado ---
    st.markdown("<div style='margin-top:72px;'></div>", unsafe_allow_html=True)
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
    # Governance Violations = solo findings con violation flag Y status open.
    # Un finding cerrado ya no es una violación activa.
    total_violations = sum(1 for f in all_findings if f.get("governance_violation") and f.get("status") == "open")
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
        viol_d = sum(1 for f in decisions if f.get("governance_violation") and f.get("status") == "open")
        st.markdown("**🛡️ Decisions**")
        st.markdown(f"- Total: {len(decisions)}")
        st.markdown(f"- Open: {open_d}")
        st.markdown(f"- Active Violations: {viol_d}")

    with col2:
        changes = [f for f in all_findings if f.get("category") == "change"]
        open_c = sum(1 for f in changes if f.get("status") == "open")
        viol_c = sum(1 for f in changes if f.get("governance_violation") and f.get("status") == "open")
        st.markdown("**🔄 Changes**")
        st.markdown(f"- Total: {len(changes)}")
        st.markdown(f"- Open: {open_c}")
        st.markdown(f"- Active Violations: {viol_c}")

    with col3:
        risks = [f for f in all_findings if f.get("category") == "risk"]
        open_r = sum(1 for f in risks if f.get("status") == "open")
        viol_r = sum(1 for f in risks if f.get("governance_violation") and f.get("status") == "open")
        st.markdown("**⚠️ Risks**")
        st.markdown(f"- Total: {len(risks)}")
        st.markdown(f"- Open: {open_r}")
        st.markdown(f"- Active Violations: {viol_r}")

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
                st.session_state["analysis_running"] = False
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
                    # Doble Vía: raw para Claude (semántica intacta),
                    # normalized para micro-triggers (keyword matching)
                    raw_text, normalized_text = extract_all_text(uploaded_files)
                    char_count = len(raw_text)

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

                        # Chunking semántico — divide en párrafos naturales con overlap
                        # Reemplaza el truncado rígido raw_text[:MAX_CHARS]
                        chunks = chunk_text_smart(raw_text, max_chars=MAX_CHARS, overlap_chars=500)
                        total_chunks = len(chunks)

                        if total_chunks > 1:
                            st.info(f"📄 Document set exceeds {MAX_CHARS:,} chars — "
                                    f"analyzing in {total_chunks} semantic chunks with context overlap.")

                        # Analizar chunk por chunk — combinar resultados
                        all_responses = []
                        for chunk_idx, chunk in enumerate(chunks, 1):
                            spinner_msg = (
                                f"🔍 Scanning chunk {chunk_idx} of {total_chunks}..."
                                if total_chunks > 1
                                else "🔍 Stellae is scanning your documents -- this may take 20-40 seconds..."
                            )
                            with st.spinner(spinner_msg):
                                chunk_response = run_anthropic_analysis(chunk, project_context=project_ctx)
                                all_responses.append(chunk_response)

                        # Combinar respuestas de todos los chunks en un único texto
                        # parse_and_store_results procesa el texto combinado
                        response_text = "\n".join(all_responses)
                        char_count = min(char_count, MAX_CHARS * total_chunks)

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
                # Limpiar cache de session_state para que recargue desde Supabase al rerun
                # NO modificar directamente — causa error "cannot be modified after widget instantiated"
                for _k in [f"ctx_ind_{project_id}", f"ctx_type_{project_id}",
                           f"ctx_stg_{project_id}", f"ctx_desc_{project_id}"]:
                    if _k in st.session_state:
                        del st.session_state[_k]
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
                            "trigger_keywords": [normalize_text(kw) for kw in rule.get("trigger_keywords", []) if kw],
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

    # ── Filtro de findings por status ────────────────────────────────────────
    # findings no tiene project_id directo — se accede via analysis_id
    _analysis_ids = [a["id"] for a in analyses_response.data]
    _all_proj_findings_resp = (
        supabase_client.table("findings")
        .select("id, status, governance_violation")
        .in_("analysis_id", _analysis_ids)
        .execute()
    )
    _all_proj_findings = _all_proj_findings_resp.data or []
    _count_open     = sum(1 for f in _all_proj_findings if f.get("status") == "open")
    _count_review   = sum(1 for f in _all_proj_findings if f.get("status") == "in_review")
    _count_closed   = sum(1 for f in _all_proj_findings if f.get("status") == "closed")
    _count_all      = len(_all_proj_findings)

    _tab_all, _tab_open, _tab_review, _tab_closed = st.tabs([
        f"All ({_count_all})",
        f"⚠️ Open ({_count_open})",
        f"🔄 In Review ({_count_review})",
        f"✅ Closed ({_count_closed})",
    ])

    def _render_analyses_filtered(status_filter):
        """Renderiza el loop de análisis mostrando solo findings con el status indicado.
        status_filter=None muestra todos."""
        # Sufijo único por tab para evitar StreamlitDuplicateElementKey
        _tab_id = status_filter if status_filter else "all"

        for analysis in analyses_response.data:
            date_label = format_analysis_date(analysis.get("created_at", ""))
            docs = analysis.get("documents_analyzed", [])
            chars = analysis.get("characters_processed", 0)
            analysis_id = analysis["id"]

            # Cargar findings de este análisis
            _f_resp = (
                supabase_client.table("findings")
                .select("*")
                .eq("analysis_id", analysis_id)
                .execute()
            )
            findings_all = _f_resp.data or []

            # Filtrar según el tab activo
            if status_filter is None:
                findings_visible = findings_all
            else:
                findings_visible = [f for f in findings_all if f.get("status") == status_filter]

            # Ocultar el expander completo si no hay findings visibles en este filtro
            if status_filter is not None and not findings_visible:
                continue

            viol_count = sum(1 for f in findings_all if f.get("governance_violation") and f.get("status") == "open")
            viol_label = f"🚨 {viol_count} active violations" if viol_count > 0 else "✅ No active violations"
            filtered_label = f" — {len(findings_visible)} {status_filter}" if status_filter else ""

            with st.expander(f"📅 {date_label} — {len(docs)} document(s) — {viol_label}{filtered_label}"):
                col_info, col_export, col_delete = st.columns([4, 2, 1])
                with col_info:
                    st.markdown("**Documents analyzed:**")
                    for doc_name in docs:
                        st.markdown(f"- {doc_name}")
                    st.markdown(f"**Characters processed:** {chars:,}")
                with col_export:
                    st.markdown("&nbsp;", unsafe_allow_html=True)
                    _audit_lang_options = list(OUTPUT_LANGUAGES.keys())
                    _audit_lang_key = f"audit_lang_{analysis_id}_{_tab_id}"
                    _audit_lang_sel = st.selectbox(
                        "Language",
                        options=_audit_lang_options,
                        index=_audit_lang_options.index(
                            st.session_state.get("pdf_lang_selected", "English")
                        ),
                        key=_audit_lang_key,
                        label_visibility="collapsed",
                    )
                    if st.button("📄 Export PDF", key=f"export_analysis_{analysis_id}_{_tab_id}",
                                 use_container_width=True):
                        if findings_all:
                            _pdf_lang = OUTPUT_LANGUAGES[_audit_lang_sel]
                            with st.spinner("Generating PDF..."):
                                try:
                                    _pdf_findings = translate_findings_for_pdf(findings_all, _pdf_lang)
                                    _open   = sum(1 for f in findings_all if f.get("status") == "open")
                                    _violations = sum(1 for f in findings_all if f.get("governance_violation") and f.get("status") == "open")
                                    if _violations > 0:
                                        _slabel, _smsg = "CRITICAL", f"{_violations} governance violation(s) detected"
                                    elif _open > 0:
                                        _slabel, _smsg = "AT RISK", f"{_open} open finding(s)"
                                    else:
                                        _slabel, _smsg = "UNDER CONTROL", "No critical issues detected"
                                    _pdf_bytes = generate_executive_pdf(
                                        selected_name, _slabel, _smsg,
                                        _pdf_findings, [analysis],
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
                                        key=f"dl_analysis_{analysis_id}_{_tab_id}",
                                        use_container_width=True,
                                    )
                                except Exception as e:
                                    st.error(f"❌ PDF error: {e}")
                        else:
                            st.warning("No findings found for this analysis.")

                with col_delete:
                    st.markdown("&nbsp;", unsafe_allow_html=True)
                    if st.button("🗑️", key=f"del_analysis_{analysis_id}_{_tab_id}",
                                 type="secondary", use_container_width=True,
                                 help="Delete this analysis"):
                        st.session_state.confirm_delete_analysis = analysis_id
                        st.session_state.confirm_delete_analysis_date = date_label
                        st.rerun()

                if st.session_state.get("confirm_delete_analysis") == analysis_id:
                    st.warning(
                        f"⚠️ Delete analysis from **{date_label}**? "
                        f"This will permanently remove all {len(docs)} document(s) and all their findings. "
                        f"Dashboard metrics will update automatically."
                    )
                    col_yes, col_no = st.columns([1, 1])
                    with col_yes:
                        if st.button("✅ Yes, delete permanently",
                                     key=f"confirm_del_{analysis_id}_{_tab_id}",
                                     type="primary", use_container_width=True):
                            try:
                                supabase_client.table("analyses").delete().eq("id", analysis_id).execute()
                                del st.session_state["confirm_delete_analysis"]
                                del st.session_state["confirm_delete_analysis_date"]
                                st.toast("✅ Analysis deleted — metrics updated.")
                                st.rerun()
                            except Exception as e:
                                st.error(f"❌ Failed to delete: {e}")
                    with col_no:
                        if st.button("❌ Cancel",
                                     key=f"cancel_del_{analysis_id}_{_tab_id}",
                                     use_container_width=True):
                            del st.session_state["confirm_delete_analysis"]
                            del st.session_state["confirm_delete_analysis_date"]
                            st.rerun()

                st.markdown("---")
                if status_filter is None:
                    st.markdown("**Findings — edit status below:**")
                    render_findings_editor(supabase_client, analysis_id)
                else:
                    # Vista de solo lectura con lista filtrada por status
                    st.markdown(f"**Findings with status: {status_filter.replace('_', ' ').title()}**")
                    if findings_visible:
                        for f in findings_visible:
                            _cat = f.get("category", "").upper()
                            _viol = "🚨 " if f.get("governance_violation") else ""
                            _title = f.get("title") or f.get("violated_rule") or "Finding"
                            with st.expander(f"{_viol}[{_cat}] {_title}"):
                                st.markdown(f.get("content", ""))
                                if f.get("action_taken"):
                                    st.caption(f"✅ Action taken: {f['action_taken']}")
                                if f.get("responsible"):
                                    st.caption(f"👤 Responsible: {f['responsible']}")
                    else:
                        st.caption("No findings with this status in this analysis.")

    with _tab_all:
        _render_analyses_filtered(None)

    with _tab_open:
        if _count_open == 0:
            st.info("✅ No open findings for this project.")
        else:
            _render_analyses_filtered("open")

    with _tab_review:
        if _count_review == 0:
            st.info("No findings currently in review.")
        else:
            _render_analyses_filtered("in_review")

    with _tab_closed:
        if _count_closed == 0:
            st.info("No closed findings yet.")
        else:
            _render_analyses_filtered("closed")


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

