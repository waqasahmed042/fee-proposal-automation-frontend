"""
app.py — Concept Engineers Fee Proposal + DocuSeal Integration
Run: python app.py  →  open http://localhost:5000
"""

import os
import tempfile
import base64
from datetime import date
from pathlib import Path

from flask import Flask, request, send_file, render_template, jsonify, redirect
from docxtpl import DocxTemplate
import requests

HERE = Path(__file__).parent
TEMPLATE = HERE / "master_proposal_template.docx"

app = Flask(__name__)

try:
    from dotenv import load_dotenv
    load_dotenv()
    print("✅ .env loaded")
except ImportError:
    pass

DOCUSEAL_BASE_URL = os.getenv("DOCUSEAL_BASE_URL", "https://api.docuseal.com")
DOCUSEAL_API_KEY  = os.getenv("DOCUSEAL_API_KEY", "")
DOCUSEAL_ROLE     = os.getenv("DOCUSEAL_ROLE", "First Party")

if not DOCUSEAL_API_KEY:
    print("❌ DOCUSEAL_API_KEY missing in .env")
else:
    print(f"✅ DocuSeal ready — {DOCUSEAL_BASE_URL}")


# Helpers
def _headers():
    return {
        "X-Auth-Token": DOCUSEAL_API_KEY,
        "Content-Type": "application/json",
        "Accept": "application/json",
    }

def build_context(form) -> dict:
    def money(s):
        try:
            return float(str(s).replace(",", "").replace("$", ""))
        except Exception:
            return 0.0

    descs   = form.getlist("fee_desc[]")
    amounts = form.getlist("fee_amount[]")
    fee_items = [{"description": d.strip(), "amount": money(a)}
                 for d, a in zip(descs, amounts) if d.strip()]
    fee_total = sum(i["amount"] for i in fee_items)

    roles = form.getlist("hourly_role[]")
    rates = form.getlist("hourly_rate[]")
    hourly_items = [{"role": r.strip(), "rate": money(rt)}
                    for r, rt in zip(roles, rates) if r.strip()]

    assumptions = [a.strip() for a in form.getlist("assumption[]") if a.strip()]
    exclusions  = [e.strip() for e in form.getlist("exclusion[]")  if e.strip()]

    scope_parts = []
    if "is_da" in form:        scope_parts.append("Development Approval (DA)")
    if "is_dd" in form:        scope_parts.append("Detailed Design (DD)")
    if "is_uu" in form:        scope_parts.append("UU Endorsed Consultant")
    if "is_uw" in form:        scope_parts.append(f"UW {form.get('uw_level','Minor')} Certifier")
    if "uu_works_only" in form: scope_parts = ["Utility Works (UU Only)"]

    return {
        "firm_name":    "Concept Engineers",
        "firm_abn":     "00 000 000 000",
        "firm_web":     "conceptengineers.com.au",
        "sender_name":  form.get("sender_name",  "Gina Harper"),
        "sender_title": form.get("sender_title", "Director"),
        "sender_email": form.get("sender_email", "gina@conceptengineers.com.au"),
        "sender_phone": form.get("sender_phone", "+61 7 0000 0000"),
        "document_type":  form.get("document_type",  "Fee Proposal"),
        "proposal_ref":   form.get("proposal_ref",   ""),
        "revision":       form.get("revision",       "Rev A"),
        "proposal_date":  form.get("proposal_date",  date.today().strftime("%d %B %Y")),
        "contact_name":   form.get("contact_name",   ""),
        "company_name":   form.get("company_name",   ""),
        "project_address":form.get("project_address",""),
        "contact_email":  form.get("contact_email",  ""),
        "project_understanding": form.get("project_understanding", ""),
        "project_scope":  ", ".join(scope_parts) or "To be confirmed",
        "is_da":          "is_da"         in form,
        "is_dd":          "is_dd"         in form,
        "is_uu":          "is_uu"         in form,
        "is_uw":          "is_uw"         in form,
        "uw_level":       form.get("uw_level", "Minor"),
        "uu_works_only":  "uu_works_only" in form,
        "capability_profile": form.get("capability_profile", "residential"),
        "known_profiles": {"residential","commercial","infrastructure","stormwater"},
        "capability_title": form.get("capability_title",""),
        "capability_body":  form.get("capability_body",""),
        "fee_items":    fee_items,
        "fee_total":    fee_total,
        "show_hourly":  bool(hourly_items),
        "hourly_items": hourly_items,
        "payment_terms": form.get("payment_terms",
                         "Fees are invoiced at agreed milestones, payable within 14 days."),
        "assumptions": assumptions,
        "exclusions":  exclusions,
    }


# Pages
@app.route("/")
def index():
    return render_template("form.html", today=date.today().strftime("%d %B %Y"))


@app.route("/templates")
def templates_page():
    """Templates management page."""
    return render_template("templates.html")


@app.route("/sign/<slug>")
def sign_redirect(slug):
    """Redirect /sign/<slug> → DocuSeal signing page."""
    return redirect(f"https://docuseal.com/s/{slug}")


# Document generation
@app.route("/generate", methods=["POST"])
def generate():
    ctx = build_context(request.form)
    tpl = DocxTemplate(str(TEMPLATE))
    tpl.render(ctx)
    ref = ctx["proposal_ref"] or "proposal"

    tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
    tpl.save(tmp.name)
    tmp.close()

    # Return base64 for DocuSeal flow
    if request.args.get("base64") == "true":
        with open(tmp.name, "rb") as f:
            b64 = base64.b64encode(f.read()).decode()
        os.unlink(tmp.name)
        return jsonify({"success": True, "documentBase64": b64, "filename": f"{ref}.docx"})

    output_fmt = request.form.get("output_fmt", "docx")
    if output_fmt == "pdf":
        try:
            from docx2pdf import convert
            pdf = tmp.name.replace(".docx", ".pdf")
            convert(tmp.name, pdf)
            os.unlink(tmp.name)
            return send_file(pdf, as_attachment=True, download_name=f"{ref}.pdf",
                             mimetype="application/pdf")
        except Exception:
            pass  # fall through to docx

    return send_file(tmp.name, as_attachment=True, download_name=f"{ref}.docx",
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


# DocuSeal API routes
@app.route("/api/upload-template", methods=["POST"])
def upload_template():
    """STEP 1 — Upload filled DOCX to DocuSeal, returns template_id."""
    if not DOCUSEAL_API_KEY:
        return jsonify({"error": "DOCUSEAL_API_KEY not set"}), 500

    data = request.json or {}
    b64 = data.get("documentBase64")
    name = data.get("templateName", "Fee Proposal")

    if not b64:
        return jsonify({"error": "documentBase64 is required"}), 400

    try:
        resp = requests.post(
            f"{DOCUSEAL_BASE_URL}/templates/docx",
            json={"name": name, "documents": [{"name": "proposal.docx", "file": b64}]},
            headers=_headers(),
            timeout=30,
        )
        resp.raise_for_status()
        result = resp.json()
        template_id = result.get("id")

        if not template_id:
            return (
                jsonify(
                    {
                        "error": "DocuSeal did not return a template ID",
                        "details": result,
                    }
                ),
                500,
            )

        # Auto-add signature field if template has none
        fields = result.get("fields", [])
        has_sig = any(f.get("name") == "signature" for f in fields)
        if not has_sig:
            schema = result.get("schema", [])
            att_uuid = schema[0].get("attachment_uuid") if schema else None
            if att_uuid:
                requests.put(
                    f"{DOCUSEAL_BASE_URL}/templates/{template_id}",
                    json={
                        "fields": [
                            {
                                "name": "signature",
                                "role": "First Party",
                                "type": "signature",
                                "required": True,
                                "areas": [
                                    {
                                        "x": 0.05,
                                        "y": 0.88,
                                        "w": 0.25,
                                        "h": 0.04,
                                        "page": 2,
                                        "attachment_uuid": att_uuid,
                                    }
                                ],
                            },
                            {
                                "name": "name",
                                "role": "First Party",
                                "type": "text",
                                "required": True,
                                "areas": [
                                    {
                                        "x": 0.55,
                                        "y": 0.88,
                                        "w": 0.25,
                                        "h": 0.03,
                                        "page": 2,
                                        "attachment_uuid": att_uuid,
                                    }
                                ],
                            },
                            {
                                "name": "position",
                                "role": "First Party",
                                "type": "text",
                                "required": True,
                                "areas": [
                                    {
                                        "x": 0.05,
                                        "y": 0.93,
                                        "w": 0.25,
                                        "h": 0.03,
                                        "page": 2,
                                        "attachment_uuid": att_uuid,
                                    }
                                ],
                            },
                            {
                                "name": "date",
                                "role": "First Party",
                                "type": "date",
                                "required": True,
                                "areas": [
                                    {
                                        "x": 0.55,
                                        "y": 0.93,
                                        "w": 0.20,
                                        "h": 0.03,
                                        "page": 2,
                                        "attachment_uuid": att_uuid,
                                    }
                                ],
                            },
                        ]
                    },
                    headers=_headers(),
                    timeout=15,
                )

        print(f"[upload] template_id={template_id} name={name}")
        return jsonify(
            {
                "success": True,
                "templateId": template_id,
                "templateName": result.get("name"),
            }
        )

    except requests.HTTPError as e:
        body = {}
        try:
            body = e.response.json()
        except Exception:
            pass
        return (
            jsonify({"error": body.get("error") or str(e), "details": body}),
            e.response.status_code,
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/get-template/<int:template_id>", methods=["GET"])
def get_template(template_id):
    """STEP 2 — Confirm template exists, return its roles."""
    if not DOCUSEAL_API_KEY:
        return jsonify({"error": "DOCUSEAL_API_KEY not set"}), 500
    try:
        resp = requests.get(f"{DOCUSEAL_BASE_URL}/templates/{template_id}",
                            headers=_headers(), timeout=10)
        resp.raise_for_status()
        result = resp.json()
        submitters = result.get("submitters", [])
        roles = [s.get("name") for s in submitters if s.get("name")]
        return jsonify({"success": True, "templateId": result.get("id"),
                        "templateName": result.get("name"), "roles": roles})
    except requests.HTTPError as e:
        body = {}
        try: body = e.response.json()
        except Exception: pass
        return jsonify({"error": body.get("error") or str(e)}), e.response.status_code
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/send-proposal", methods=["POST"])
def send_proposal():
    """
    STEP 3 — Create DocuSeal submission.
    Dynamically resolves role from template to avoid 422 errors.
    Returns signing_url for client.
    """
    if not DOCUSEAL_API_KEY:
        return jsonify({"error": "DOCUSEAL_API_KEY not set"}), 500

    data          = request.json or {}
    template_id   = data.get("templateId")
    client_name   = (data.get("client_name")   or "").strip()
    client_email  = (data.get("client_email")  or "").strip().lower()
    project_addr  = (data.get("project_address") or "").strip()

    if not template_id:   return jsonify({"error": "templateId is required"}), 400
    if not client_name:   return jsonify({"error": "client_name is required"}), 400
    if not client_email:  return jsonify({"error": "client_email is required"}), 400
    if not project_addr:  return jsonify({"error": "project_address is required"}), 400

    # Resolve real role + field names from template
    role = DOCUSEAL_ROLE
    template_fields = []
    try:
        tr = requests.get(f"{DOCUSEAL_BASE_URL}/templates/{template_id}",
                          headers=_headers(), timeout=10)
        tr.raise_for_status()
        td = tr.json()
        submitters = td.get("submitters", [])
        if submitters:
            role = submitters[0].get("name", role)
        template_fields = [f.get("name") for f in td.get("fields", []) if f.get("name")]
        print(f"[send] role='{role}' fields={template_fields}")
    except Exception as e:
        print(f"[send] template fetch failed, using defaults: {e}")

    # Only prefill fields that actually exist in template
    prefill = {"client_name": client_name, "project_address": project_addr}
    fields_payload = [{"name": k, "default_value": v}
                      for k, v in prefill.items() if k in template_fields]

    submitter = {"role": role, "name": client_name, "email": client_email}
    if fields_payload:
        submitter["fields"] = fields_payload

    payload = {"template_id": int(template_id), "send_email": True,
               "submitters": [submitter]}

    print(f"[send] payload={payload}")

    try:
        resp = requests.post(f"{DOCUSEAL_BASE_URL}/submissions",
                             json=payload, headers=_headers(), timeout=20)
        print(f"[send] status={resp.status_code} body={resp.text[:400]}")
        resp.raise_for_status()

        result = resp.json()
        sub = result[0] if isinstance(result, list) else result.get("submitters", [{}])[0]
        slug = sub.get("slug")
        signing_url = f"https://docuseal.com/s/{slug}" if slug else None

        print(f"[send] ✅ sent to {client_email} signingUrl={signing_url}")
        return jsonify({"success": True, "signingUrl": signing_url,
                        "submissionId": sub.get("submission_id"),
                        "message": f"Proposal sent to {client_email}"})

    except requests.HTTPError as e:
        body = {}
        try: body = e.response.json()
        except Exception: pass
        msg = body.get("error") or body.get("message") or str(e)
        print(f"[send] HTTP {e.response.status_code}: {body}")
        return jsonify({"error": msg, "details": body}), e.response.status_code
    except Exception as e:
        print(f"[send] Exception: {e}")
        return jsonify({"error": str(e)}), 500


@app.route("/api/all-templates", methods=["GET"])
def all_templates():
    """Get all templates from DocuSeal."""
    if not DOCUSEAL_API_KEY:
        return jsonify({"error": "DOCUSEAL_API_KEY not set"}), 500
    try:
        resp = requests.get(f"{DOCUSEAL_BASE_URL}/templates",
                            headers=_headers(), timeout=15)
        resp.raise_for_status()
        data = resp.json()
        # DocuSeal returns { data: [...] } or plain list
        templates = data.get("data", data) if isinstance(data, dict) else data
        print(f"[all-templates] fetched {len(templates)} templates")
        return jsonify({"success": True, "templates": templates, "count": len(templates)})
    except requests.HTTPError as e:
        body = {}
        try: body = e.response.json()
        except Exception: pass
        return jsonify({"error": body.get("error") or str(e)}), e.response.status_code
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/delete-template/<int:template_id>", methods=["DELETE"])
def delete_template(template_id):
    """Delete a template from DocuSeal."""
    if not DOCUSEAL_API_KEY:
        return jsonify({"error": "DOCUSEAL_API_KEY not set"}), 500
    try:
        resp = requests.delete(f"{DOCUSEAL_BASE_URL}/templates/{template_id}",
                               headers=_headers(), timeout=10)
        resp.raise_for_status()
        return jsonify({"success": True, "message": f"Template {template_id} deleted"})
    except requests.HTTPError as e:
        body = {}
        try: body = e.response.json()
        except Exception: pass
        return jsonify({"error": body.get("error") or str(e)}), e.response.status_code
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(debug=True, port=5000)
