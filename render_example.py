"""
render_example.py
Shows how the generation engine fills the master template for a given proposal
configuration. In production, `context` is built from the staff tick-box form.
"""

from pathlib import Path

from docxtpl import DocxTemplate, RichText

HERE = Path(__file__).parent
TEMPLATE = str(HERE / "master_proposal_template.docx")
KNOWN_PROFILES = {"residential", "commercial", "infrastructure", "stormwater"}


def build_context(cfg: dict, tpl: DocxTemplate) -> dict:
    """Translate a staff selection dict into a docxtpl context."""
    fee_items = cfg["fee_items"]
    fee_total = sum(i["amount"] for i in fee_items)

    return {
        # firm / sender
        "firm_name": "Concept Engineers",
        "firm_abn": "00 000 000 000",
        "firm_web": "conceptengineers.com.au",
        "sender_name": cfg.get("sender_name", "Gina Harper"),
        "sender_title": cfg.get("sender_title", "Director"),
        "sender_email": cfg.get("sender_email", "gina@conceptengineers.com.au"),
        "sender_phone": cfg.get("sender_phone", "+61 7 0000 0000"),
        # document meta
        "document_type": cfg.get("document_type", "Fee Proposal"),
        "proposal_ref": cfg["proposal_ref"],
        "revision": cfg.get("revision", "Rev A"),
        "proposal_date": cfg["proposal_date"],
        # client
        "contact_name": cfg["contact_name"],
        "company_name": cfg["company_name"],
        "project_address": cfg["project_address"],
        "contact_email": cfg["contact_email"],
        "project_understanding": cfg["project_understanding"],
        # scope flags
        "is_da": cfg.get("is_da", False),
        "is_dd": cfg.get("is_dd", False),
        "is_uu": cfg.get("is_uu", False),
        "is_uw": cfg.get("is_uw", False),
        "uw_level": cfg.get("uw_level", "Minor"),
        "uu_works_only": cfg.get("uu_works_only", False),
        # capability page (1 of 9)
        "capability_profile": cfg.get("capability_profile", "residential"),
        "known_profiles": KNOWN_PROFILES,
        "capability_title": cfg.get("capability_title", ""),
        "capability_body": cfg.get("capability_body", ""),
        # fees
        "fee_items": fee_items,
        "fee_total": fee_total,
        "show_hourly": bool(cfg.get("hourly_items")),
        "hourly_items": cfg.get("hourly_items", []),
        "payment_terms": cfg.get("payment_terms",
                                 "Fees are invoiced at agreed milestones, payable within 14 days."),
        # assumptions / exclusions
        "assumptions": cfg.get("assumptions", []),
        "exclusions": cfg.get("exclusions", []),
    }


def render(cfg, out_path):
    tpl = DocxTemplate(TEMPLATE)
    tpl.render(build_context(cfg, tpl))
    tpl.save(out_path)
    print("Rendered:", out_path)


if __name__ == "__main__":
    base = dict(
        proposal_ref="CE-2026-0412",
        proposal_date="15 June 2026",
        contact_name="David Lee",
        company_name="Skyline Developments Pty Ltd",
        project_address="42 Riverside Drive, Brisbane QLD 4000",
        contact_email="david@skyline.com.au",
        project_understanding=(
            "We understand the project involves a residential development on the "
            "subject site requiring civil engineering input to support approval and "
            "delivery. Our scope is tailored to the components selected below."),
        terms_url="https://conceptengineers.com.au/terms",
        assumptions=[
            "A current detail survey will be provided by the client.",
            "Geotechnical information will be made available where required.",
        ],
        exclusions=[
            "Authority application and lodgement fees.",
            "Structural engineering of buildings.",
        ],
    )

    # Scenario A: DA + DD, full scope, residential capability
    render({**base,
            "is_da": True, "is_dd": True,
            "capability_profile": "residential",
            "fee_items": [
                {"description": "Development Approval (DA) civil engineering", "amount": 8500.00},
                {"description": "Detailed Design (DD) and documentation", "amount": 14200.00},
            ],
            "hourly_items": [
                {"role": "Principal Engineer", "rate": 240.00},
                {"role": "Senior Engineer", "rate": 185.00},
            ]},
           str(HERE / "sample_A_da_dd.docx"))

    # Scenario B: DA only, commercial capability, no hourly
    render({**base,
            "proposal_ref": "CE-2026-0413",
            "is_da": True,
            "capability_profile": "commercial",
            "fee_items": [
                {"description": "Development Approval (DA) civil engineering", "amount": 9200.00},
            ]},
           str(HERE / "sample_B_da_only.docx"))

    # Scenario C: UU works only, simplified
    render({**base,
            "proposal_ref": "CE-2026-0414",
            "is_uu": True, "uu_works_only": True,
            "capability_profile": "stormwater",
            "fee_items": [
                {"description": "Urban Utilities water & sewer design (works only)", "amount": 6400.00},
            ]},
           str(HERE / "sample_C_uu_only.docx"))
