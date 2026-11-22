# openUI Standard Consultant Profile

## Purpose

Standardize consultant information for registration, review, matching, public profiles, and later CV exports. The structured website record is authoritative. An uploaded CV is private supporting evidence.

The structure is inspired by the useful project-experience approach found in the supplied ADB-style CV, but it is an openUI format and must not be described as ADB-approved.

## Representations

Maintain three views from the same approved structured data:

| View | Contents |
| --- | --- |
| Private record | Full approved profile, contact data, private employment/projects, CV file, references, verification and audit history |
| Public profile | Consultant-approved and administrator-approved professional fields only |
| Generated CV | Later export containing approved fields selected for a general or assignment-specific CV |

## Consultant form sections

### Account and identity

- Legal name: private, required.
- Display name: public-eligible, required.
- Email and phone: private, required as configured for account operation.
- Country of residence and time zone: private by default; region/time-zone summary may be public.
- Profile photo: optional; must not influence matching.

### Professional summary

- Professional headline: required, public-eligible.
- Concise professional summary: required, public-eligible.
- Primary expertise categories: required, controlled vocabulary.
- Secondary expertise and specialist capabilities: optional, controlled vocabulary plus reviewed additions.
- Years of relevant experience: calculated or administrator-approved; do not infer age.

### Languages and regional experience

- Language, proficiency, and verification status where applicable.
- Country/region, type of experience, and approximate duration.
- Travel or work-location availability where relevant.

### Education and credentials

Use repeatable records:

- Institution or issuing body.
- Qualification, degree, certification, membership, or training title.
- Field or specialization.
- Completion/award year.
- Credential identifier and expiry where relevant, private by default.
- Verification status and evidence reference, private.
- Public-visibility choice.

### Employment and consultancy assignments

Use repeatable reverse-chronological records:

- Organization or client name.
- Role or position.
- Engagement type.
- Start and end month/year; allow `present`.
- Country/region and remote/on-site status.
- Client sector.
- Short organization/project context.
- Responsibilities.
- Deliverables.
- Measurable achievements or outcomes.
- Expertise demonstrated.
- Technologies, standards, or methods used.
- Reference availability and private reference details.
- Public-visibility choice.

### Projects

Use a separate repeatable project record when one assignment contains several material projects:

- Project title and organization/client.
- Consultant role.
- Start/end dates and location/region.
- Project context and objective.
- Responsibilities and deliverables.
- Measurable achievements.
- Expertise, technologies, and methods demonstrated.
- Client sector and reference availability.
- Link to the related employment/consultancy record.
- Public-visibility choice.

### Availability and commercial status

- Working time zone.
- Normal availability windows.
- Earliest start date.
- Supported consultation languages.
- Approved consultant tier and effective date.
- Tier approval and consultant acceptance status.

Consultants may view the tier but may not approve their own tier.

### Declaration and workflow

- Consultant declaration that entered information is accurate.
- Consent for private CV storage and later proposal-only extraction.
- Consent for selected fields to appear publicly.
- Draft, submitted, needs-correction, under-review, approved, suspended, withdrawn, and archived states.
- Submitter, reviewer, timestamps, version, and audit history.

## Excluded or restricted fields

Do not request or infer for normal registration or matching:

- Date of birth or age.
- Health, disability, height, weight, or medical information.
- Marital status, family status, spouse, or children.
- Full residential address.
- Religion, ethnicity, political views, or similar sensitive traits.
- Personal circumstances unrelated to professional suitability.

Citizenship/nationality may be added only for an approved operational or legal requirement. It remains private and must not affect matching unless the requirement explicitly and lawfully demands it.

## Public profile rules

Public fields require both consultant visibility selection and administrator approval. Never publish:

- Legal identity or private contact details.
- Original CV or private file links.
- References or verification evidence.
- Unapproved organizations, clients, projects, credentials, or claims.
- Private employment/project records.

Public profiles may show display name, headline, summary, approved expertise, languages, regional experience, selected assignments/projects, approved education/credentials, tier/rate, and availability summary.

## Matching rules

Matching may use only approved fields marked as matching-eligible, such as:

- Expertise and demonstrated project evidence.
- Languages.
- Availability and time-zone compatibility.
- Country/regional experience where relevant to the request.
- Sector, technology, qualification, and role experience.
- Approved tier and client price.

Do not use raw CV text, unapproved extraction proposals, profile photo, protected traits, private contact data, or irrelevant personal data.

## Manual-first and later extraction

For the initial consultant-registration phase:

1. Consultant completes the structured form manually.
2. Consultant uploads the original CV privately.
3. Administrator compares the record with the CV and requests corrections.
4. Consultant confirms corrections.
5. Administrator approves the profile and eligible public fields.

In the later n8n phase, extraction may propose schema-aligned additions/corrections, missing fields, conflicts, expertise tags, and a tier suggestion. Every proposal must retain source evidence and confidence and requires human review; it never overwrites approved data.

## Validation and completion criteria

- Required fields and controlled vocabularies validate on client and server.
- Date ranges are coherent and overlapping assignments are allowed but flagged for review.
- `present` is permitted only for active records.
- Public/private/matching visibility is explicit and enforced server-side.
- Every approved field version has provenance and reviewer history.
- CV replacement does not silently change approved structured data.
- Synthetic tests cover long careers, simultaneous roles, incomplete dates, multiple languages/regions, disputed extraction, withdrawn consent, and public/private separation.
- A consultant can complete, save, review, submit, correct, and obtain approval without n8n or AI.
