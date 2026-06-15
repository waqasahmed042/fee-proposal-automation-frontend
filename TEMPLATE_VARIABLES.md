# Master Proposal Template — Variable Reference

This is the data contract between the staff tick-box form and the document
engine (`docxtpl`). The form produces these values; the engine renders the
template into a finished proposal.

## How the conditional logic works

The template uses three docxtpl tag types. This is what delivers the
"no blank pages, no layout breaks" requirement:

- `{{ variable }}` — inline value substitution.
- `{%p if flag %} ... {%p endif %}` — the `p` prefix deletes the **whole
  paragraph** carrying the tag, so a hidden section leaves nothing behind.
- `{%tr for x in list %}` / `{%tr endfor %}` — placed on their **own rows**,
  each deletes its whole table row, so fee/hourly rows repeat or vanish cleanly.

Never put `for` and `endfor` in the same table row — docxtpl 0.20.x will drop
the opening tag. Keep them on separate control rows around the body row.

## Variables

### Firm / sender (usually constant)
| Variable | Example |
|---|---|
| `firm_name` | Concept Engineers |
| `firm_abn` | 00 000 000 000 |
| `firm_web` | conceptengineers.com.au |
| `sender_name` | Gina Harper |
| `sender_title` | Director |
| `sender_email` | gina@conceptengineers.com.au |
| `sender_phone` | +61 7 0000 0000 |

### Document meta
| Variable | Notes |
|---|---|
| `document_type` | e.g. "Fee Proposal". Used in title, subject, footer. |
| `proposal_ref` | Your reference, e.g. CE-2026-0412. |
| `revision` | Rev A / Rev B / Rev C. |
| `proposal_date` | Display date string. |

### Client / project
| Variable | Notes |
|---|---|
| `contact_name` | Recipient. |
| `company_name` | Client company. |
| `project_address` | Site address. |
| `contact_email` | Recipient email. |
| `project_understanding` | One short paragraph describing the project. |

### Scope flags (booleans — drive which scope sections appear)
| Variable | Shows |
|---|---|
| `is_da` | Development Approval scope block. |
| `is_dd` | Detailed Design scope block. |
| `is_uu` | Urban Utilities endorsed-consultant block. |
| `is_uw` | Certifier block; `uw_level` = "Minor" or "Major". |
| `uu_works_only` | Simplified UU-works-only note. |

### Capability page (1 of 9)
| Variable | Notes |
|---|---|
| `capability_profile` | Key selecting the capability block. Built-in keys: `residential`, `commercial`, `infrastructure`, `stormwater`. |
| `known_profiles` | Set of built-in keys (pass it so the fallback works). |
| `capability_title`, `capability_body` | Used only when `capability_profile` is not a built-in key (custom fallback). |

To add the remaining 5 profiles, copy a `{%p if capability_profile == 'key' %}`
block in `build_template.py` and add the key to `known_profiles`.

### Fees
| Variable | Type | Notes |
|---|---|---|
| `fee_items` | list of `{description, amount}` | Each becomes a fee-table row. |
| `fee_total` | number | Sum of `amount`; rendered in the total row. |
| `show_hourly` | bool | Shows/hides the entire hourly-rate block. |
| `hourly_items` | list of `{role, rate}` | Hourly-rate rows. |
| `payment_terms` | string | Milestone/payment sentence. |

### Assumptions / exclusions
| Variable | Type | Notes |
|---|---|---|
| `assumptions` | list of strings | Heading + bullets appear only if non-empty. |
| `exclusions` | list of strings | Same. |

### Terms
The standard-terms hyperlink is baked into the template (points to
`conceptengineers.com.au/terms`). To change it, edit the one `add_hyperlink(...)`
line in `build_template.py`. It is intentionally not a per-proposal variable,
since the terms page is constant.

## Files
- `master_proposal_template.docx` — the template the engine consumes.
- `build_template.py` — regenerates the template (edit this to change layout/sections).
- `render_example.py` — `build_context()` + three worked scenarios.
- `sample_A_da_dd.docx` — DA + DD, residential, with hourly rates.
- `sample_B_da_only.docx` — DA only, commercial, no hourly.
- `sample_C_uu_only.docx` — UU works only, simplified.
