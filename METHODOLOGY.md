# STELLAE — Core Predictive Methodology & Mathematical Framework

Este documento detalla el sustento científico, el modelo matemático y la lógica de orquestación algorítmica que impulsa el **Motor de Inferencia de Gobernanza Bayesiana** de Stellae (Stage 1).

---

## 1. El Fundamento Académico: El Problema del "Cold Start"

El mayor desafío al implementar sistemas predictivos de Inteligencia Artificial en proyectos de capital (EPC) es la falta de datos históricos limpios y la confidencialidad de las corporaciones (*Cold Start Problem*). 

Stellae resuelve esto operationalizando la literatura científica de gestión de megaproyectos (específicamente los marcos de Oxford Global Projects, el IPA y la investigación cuantitativa de gobernanza de Shen et al.).

### El Coeficiente de Confianza Inicial (Prior Bayesiano)
Establecemos un estado de confianza inicial fijo para la salud de gobernanza del proyecto:

$$\beta = 0.88$$

Este coeficiente ($\beta$) actúa como nuestro **Prior Bayesiano**. Representa la probabilidad base de que existan desviaciones ocultas o fallas de alineación en interfaces complejas antes de analizar la documentación viva del proyecto. El sistema no necesita aprender desde cero; asume la probabilidad histórica de la industria pesada desde el Día 1 del piloto.

---

## 2. La Innovación Tecnológica: Métricas Proxy Objetivas

La investigación académica tradicional (incluyendo a Shen) mide la salud de la gobernanza mediante **encuestas subjetivas periódicas** a los miembros del equipo. Este enfoque metodológico falla en el campo debido a:
1. Fatiga de llenado por parte del staff.
2. Sesgo de optimismo (los ingenieros tienden a reportar que "todo va bien" hasta que explota la crisis).
3. Latencia (las encuestas son mensuales; los desastres son diarios).

### El Enfoque de Stellae
Stellae sustituye las encuestas subjetivas por **Métricas Proxy Objetivas y Automatizadas** extraídas directamente del flujo de documentos contractuales semanales:

* **Latencia de Respuesta Contractual (Contractual Response Latency):** El desfase de marcas de tiempo entre la emisión de una RFI crítica y su respuesta formal.
* **Omisión de Pasos del DoA (DoA Step Omissions):** Decisiones registradas en minutas que por su impacto financiero o técnico requerían la firma de un nivel superior según la *Authority Matrix*, pero no generaron el documento de salida en las 24-72 horas estipuladas.
* **Patrones de Riesgo Normalizados (Normalized Risk Patterns):** Recurrencia de incidentes HSE idénticos o near-misses en el mismo TAG de equipo en menos de 30 días, sin reporte formal al Steering Committee.

Estas métricas proxy alimentan dinámicamente el motor, modificando probabilísticamente el prior inicial. El proxy de Stellae es técnicamente superior al usado en los papers científicos porque se basa en **hechos digitales inalterables, no en opiniones**.

---

## 3. Heurística de Decaimiento y Costo de la Inacción

### Constante de Decaimiento (Decay Constant)
El Stage 1 de Stellae implementa un factor de decaimiento lineal inicial de **0.12** por semana de inacción detectada. Si un hallazgo (*Finding*) crítico permanece en estado `Open` o sin el documento de remediación contractual, la probabilidad de sobrecoste e impacto en la ruta crítica escala exponencialmente según la fase constructiva.

*Nota del Roadmap:* Este factor de 0.12 es una heurística calibrada por la experiencia de 12 años en campo del fundador. En el Stage 2, esta constante se optimizará mediante redes neuronales tradicionales a medida que el sistema ingiera los resultados reales de los cierres de proyectos de los clientes.

### Algoritmo del Costo de la Inacción (COI)
El sistema traduce la inacción en impacto financiero directo cruzando tres variables en la base de datos (Supabase):

$$COI = (Días\ de\ Retraso\ Latente \times Burn\ Rate\ Diario) \times Multiplicador\ de\ Interferencia$$

Donde:
1. **Días de Retraso Latente:** Calculado por el motor de inferencia analizando la ruta crítica afectada por el componente (ej. Línea principal de gas).
2. **Burn Rate Diario ($/día):** Configurado en el contexto del proyecto (costo fijo de mantener la obra abierta).
3. **Multiplicador de Interferencia:** Factor que oscila entre 1.5x y 3.0x si el hallazgo implica demolición de obra civil o retrabajo físico (*rework*) en campo.

---
**Stellae — Governance Intelligence** *Confidential — Internal Methodology Document — May 2026*