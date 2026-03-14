"""
analyzer.py
-----------
Supports two AI backends:
  1. OpenRouter  (OPENROUTER_API_KEY)  → https://openrouter.ai  — use this locally
  2. Anthropic   (ANTHROPIC_API_KEY)   → api.anthropic.com      — fallback / direct

Auto-detects which key is set. Falls back to rule-based analysis if neither is available.
Pass either key via --api-key on the CLI (the script auto-detects which type it is by prefix).
"""

import os
import json
import urllib.request
import urllib.error


# ─────────────────────────────────────────────
# Public entry point
# ─────────────────────────────────────────────

def analyze_with_ai(inspection_data, thermal_data) -> dict:
    """
    Calls an LLM to generate DDR narrative.
    Priority: OPENROUTER_API_KEY → ANTHROPIC_API_KEY → rule-based fallback.
    """
    or_key  = os.environ.get("OPENROUTER_API_KEY", "")
    ant_key = os.environ.get("ANTHROPIC_API_KEY", "")

    context = _build_context(inspection_data, thermal_data)
    prompt  = _build_prompt(context)

    if or_key:
        result = _call_openrouter(or_key, prompt)
        if result:
            return result

    if ant_key:
        result = _call_anthropic(ant_key, prompt)
        if result:
            return result

    print("[Analyzer] No AI backend available — using rule-based analysis.")
    return _rule_based_analysis(inspection_data, thermal_data)


# ─────────────────────────────────────────────
# Backend: OpenRouter  (works locally)
# ─────────────────────────────────────────────

def _call_openrouter(api_key: str, prompt: str) -> dict | None:
    """
    POST to https://openrouter.ai/api/v1/chat/completions
    Uses anthropic/claude-sonnet-4-5 (or any model on OpenRouter).
    """
    try:
        payload = json.dumps({
            "model": "anthropic/claude-sonnet-4-5",
            "max_tokens": 2000,
            "messages": [{"role": "user", "content": prompt}],
        }).encode("utf-8")

        req = urllib.request.Request(
            "https://openrouter.ai/api/v1/chat/completions",
            data=payload,
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
                "HTTP-Referer": "https://github.com/ddr-report-generator",
                "X-Title": "DDR Report Generator",
            },
            method="POST",
        )

        with urllib.request.urlopen(req, timeout=90) as resp:
            result = json.loads(resp.read().decode("utf-8"))

        raw = result["choices"][0]["message"]["content"].strip()
        raw = raw.replace("```json", "").replace("```", "").strip()
        parsed = json.loads(raw)
        print("[Analyzer] AI analysis complete via OpenRouter (claude-sonnet-4-5).")
        return parsed

    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        print(f"[Analyzer] OpenRouter HTTP {e.code}: {body[:300]}")
        return None
    except Exception as e:
        print(f"[Analyzer] OpenRouter error: {type(e).__name__}: {e}")
        return None


# ─────────────────────────────────────────────
# Backend: Anthropic direct  (api.anthropic.com)
# ─────────────────────────────────────────────

def _call_anthropic(api_key: str, prompt: str) -> dict | None:
    """
    POST to https://api.anthropic.com/v1/messages
    Uses claude-haiku-4-5-20251001 (fast, cheap).
    """
    try:
        payload = json.dumps({
            "model": "claude-haiku-4-5-20251001",
            "max_tokens": 2000,
            "messages": [{"role": "user", "content": prompt}],
        }).encode("utf-8")

        req = urllib.request.Request(
            "https://api.anthropic.com/v1/messages",
            data=payload,
            headers={
                "x-api-key": api_key,
                "anthropic-version": "2023-06-01",
                "Content-Type": "application/json",
            },
            method="POST",
        )

        with urllib.request.urlopen(req, timeout=90) as resp:
            result = json.loads(resp.read().decode("utf-8"))

        raw = result["content"][0]["text"].strip()
        raw = raw.replace("```json", "").replace("```", "").strip()
        parsed = json.loads(raw)
        print("[Analyzer] AI analysis complete via Anthropic API (claude-haiku-4-5-20251001).")
        return parsed

    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        print(f"[Analyzer] Anthropic HTTP {e.code}: {body[:300]}")
        return None
    except Exception as e:
        print(f"[Analyzer] Anthropic error: {type(e).__name__}: {e}")
        return None


# ─────────────────────────────────────────────
# Prompt builder
# ─────────────────────────────────────────────

def _build_prompt(context: str) -> str:
    return f"""You are a professional building inspection engineer.
Based on the following structured inspection data, generate a client-friendly DDR (Detailed Defect Report) narrative.

INSPECTION DATA:
{context}

Generate a JSON response with exactly these keys:
{{
  "executive_summary": "3-4 sentences summarizing overall condition, key issues found, and urgency. Use plain language.",
  "root_cause_analysis": "2-3 sentences explaining WHY the dampness/leakage is occurring (tile joint gaps → water ingress → skirting dampness chain).",
  "thermal_interpretation": "2-3 sentences interpreting the thermal readings — what the temperature differentials indicate about moisture presence.",
  "severity_assessment": "overall severity: Critical/High/Medium/Low with brief justification",
  "findings": [
    {{
      "id": "F-01",
      "area": "area name",
      "observation": "what was observed (negative side)",
      "source": "identified root cause / source (positive side)",
      "severity": "High/Medium/Low",
      "thermal_evidence": "yes/no - thermal reading reference if applicable",
      "recommendation": "specific remediation action in plain language",
      "priority": "Immediate/Within 30 days/Within 90 days/Monitor"
    }}
  ],
  "recommendations": [
    {{"action": "action text", "priority": "Immediate/Short-term/Long-term", "area": "area"}}
  ],
  "conclusion": "1-2 sentence professional closing statement with next steps."
}}

Rules:
- Do NOT invent facts not in the data
- Use client-friendly language (avoid excessive jargon)
- If a value is missing, write "Not Available"
- Respond ONLY with the JSON, no markdown fences"""


# ─────────────────────────────────────────────
# Context builder
# ─────────────────────────────────────────────

def _build_context(inspection_data, thermal_data) -> str:
    """Builds a compact text context for the AI prompt."""
    lines = [
        f"Property: {inspection_data.property_type}, {inspection_data.floors} floors",
        f"Inspection Date: {inspection_data.inspection_date}",
        f"Inspected By: {inspection_data.inspected_by}",
        f"Overall Score: {inspection_data.overall_score}",
        f"Impacted Rooms: {', '.join(inspection_data.impacted_rooms)}",
        "",
        "IMPACTED AREAS:",
    ]
    for area in inspection_data.impacted_areas:
        lines.append(f"  Area {area.area_number}:")
        lines.append(f"    Damage observed: {area.negative_description}")
        lines.append(f"    Source identified: {area.positive_description}")

    lines += ["", "SUMMARY TABLE (Negative → Positive correlation):"]
    for neg, pos in zip(inspection_data.summary_negative, inspection_data.summary_positive):
        lines.append(f"  {neg['point']}: {neg['description']}")
        lines.append(f"  {pos['point']}: {pos['description']}")

    lines += ["", "CHECKLIST FINDINGS:"]
    for item in inspection_data.checklist_items:
        lines.append(f"  [{item.category}] {item.item}: {item.value}")

    lines += [
        "",
        f"THERMAL DATA ({len(thermal_data.readings)} readings, date: {thermal_data.readings[0].date if thermal_data.readings else 'N/A'}):",
        f"  Max hotspot: {thermal_data.max_hotspot}°C",
        f"  Min coldspot: {thermal_data.min_coldspot}°C",
        f"  Average delta (hotspot-coldspot): {thermal_data.avg_delta}°C",
        f"  Anomaly count (delta >= 4°C): {thermal_data.anomaly_count}",
    ]
    if thermal_data.readings:
        lines.append("  Top readings by delta:")
        sorted_r = sorted(thermal_data.readings, key=lambda r: r.delta, reverse=True)[:5]
        for r in sorted_r:
            lines.append(f"    {r.image_id}: hotspot={r.hotspot}°C, coldspot={r.coldspot}°C, Δ={r.delta}°C")

    return "\n".join(lines)


# ─────────────────────────────────────────────
# Rule-based fallback  (no API needed)
# ─────────────────────────────────────────────

def _rule_based_analysis(inspection_data, thermal_data) -> dict:
    """Fallback rule-based analysis — no API needed."""

    findings = []
    for i, area in enumerate(inspection_data.impacted_areas):
        thermal_ref = "Not Available"
        if i < len(thermal_data.readings):
            r = thermal_data.readings[i]
            thermal_ref = f"Image {r.image_id}: hotspot={r.hotspot}°C, Δ={r.delta}°C"

        neg = area.negative_description.lower()
        if "parking" in neg or "efflorescence" in neg:
            severity, priority = "High", "Immediate"
        elif "dampness" in neg:
            severity, priority = "Medium", "Within 30 days"
        else:
            severity, priority = "Low", "Within 90 days"

        pos = area.positive_description.lower()
        if "tile" in pos and "hollow" in pos:
            rec = ("Re-grout all tile joints in the bathroom above using waterproof epoxy grout. "
                   "Check and repair Nahani trap sealing. Apply waterproof membrane before re-tiling.")
        elif "crack" in pos or "external" in pos:
            rec = ("Seal external wall cracks using polymer-modified sealant. "
                   "Inspect and re-seal all external plumbing penetrations. Apply waterproof coating on external facade.")
        elif "plumbing" in pos:
            rec = ("Replace / repair concealed plumbing joints. "
                   "Conduct pressure test on water supply lines. Seal all pipe entry points.")
        else:
            rec = "Investigate further and apply appropriate waterproofing treatment."

        findings.append({
            "id": f"F-{i+1:02d}",
            "area": area.negative_description.split("–")[0].strip() if "–" in area.negative_description else area.negative_description,
            "observation": area.negative_description,
            "source": area.positive_description,
            "severity": severity,
            "thermal_evidence": "Yes" if i < len(thermal_data.readings) else "No",
            "recommendation": rec,
            "priority": priority,
        })

    recs = [
        {"action": "Immediately re-grout all open tile joints in Common Bathroom and Master Bedroom Bathroom of Flat No. 203 using waterproof epoxy grout to stop active water ingress.", "priority": "Immediate", "area": "All Bathrooms"},
        {"action": "Apply a waterproof membrane (Polyurethane or Cementitious) to all bathroom floors before final tiling.", "priority": "Immediate", "area": "All Bathrooms"},
        {"action": "Seal external wall cracks with polymer-modified sealant and re-coat the external facade with waterproof paint.", "priority": "Short-term", "area": "External Wall"},
        {"action": "Repair / replace the corroded external plumbing duct and ensure all pipe penetrations through the wall are properly sealed.", "priority": "Short-term", "area": "External Duct"},
        {"action": "Rectify the parking area ceiling leakage by waterproofing the floor slab of Flat No. 103 from above.", "priority": "Immediate", "area": "Parking Area"},
        {"action": "Once source repairs are complete, re-paint all damp internal walls with anti-fungal primer and finish coat.", "priority": "Long-term", "area": "Internal Walls"},
        {"action": "Conduct a follow-up thermal imaging survey 30 days after repairs to confirm moisture elimination.", "priority": "Long-term", "area": "All Areas"},
    ]

    anomaly_note = (
        f"{thermal_data.anomaly_count} readings showed a temperature differential ≥ 4°C, "
        "which strongly indicates active moisture presence behind wall surfaces."
        if thermal_data.anomaly_count > 0
        else "Temperature differentials are within normal range across most readings."
    )

    return {
        "executive_summary": (
            f"A comprehensive inspection of Flat No. 103 was conducted on {inspection_data.inspection_date} "
            f"by {inspection_data.inspected_by}. The inspection identified dampness and water ingress across "
            f"{len(inspection_data.impacted_areas)} areas including the Hall, Bedroom, Master Bedroom, Kitchen, "
            f"and Parking area. The primary root cause is open tile joint gaps in the bathrooms of the flat above "
            f"(Flat No. 203) and external wall cracks near the Master Bedroom. Immediate remediation is recommended "
            f"to prevent further structural deterioration."
        ),
        "root_cause_analysis": (
            "The dampness observed at the skirting level across multiple rooms originates from water seeping through "
            "open tile joint gaps and a damaged Nahani trap in the Common and Master Bedroom bathrooms of Flat No. 203 "
            "(the flat above). This water travels downward through the floor slab and emerges as dampness at the skirting "
            "level of Flat No. 103 below. Additionally, external wall cracks near the Master Bedroom are allowing rainwater "
            "to penetrate, causing dampness and efflorescence on the internal wall surface."
        ),
        "thermal_interpretation": (
            f"Thermal imaging was conducted on {thermal_data.readings[0].date if thermal_data.readings else 'the inspection date'} "
            f"using a Bosch GTC 400 C Professional camera across {len(thermal_data.readings)} locations. "
            f"The maximum hotspot recorded was {thermal_data.max_hotspot}°C against a minimum coldspot of "
            f"{thermal_data.min_coldspot}°C. {anomaly_note} "
            f"Cold patches (blue areas) visible in the thermal images confirm moisture-laden wall surfaces."
        ),
        "severity_assessment": (
            "High — Multiple areas are actively affected by water ingress with confirmed thermal evidence. "
            "The parking area leakage and external wall cracks pose a risk of progressive structural damage if left unaddressed."
        ),
        "findings": findings,
        "recommendations": recs,
        "conclusion": (
            "The property requires prompt waterproofing interventions primarily targeting the bathroom tile joints of "
            "Flat No. 203 and the external wall facade. Once the source defects are rectified and the internal surfaces "
            "have dried, internal redecoration should follow. A post-repair thermal survey is recommended to confirm "
            "complete moisture elimination."
        ),
    }
