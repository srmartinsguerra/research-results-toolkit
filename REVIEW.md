# Public-release review

Scope: the combined `RESEARCH_RESULTS.gs`, README, tests, and repository files delivered with this public package. This review did not inspect a live Google project, its Script Properties, Drive backups, or existing GitHub history.

## Findings in the previous combined source

| Item | Finding | Public-package action |
| --- | --- | --- |
| Destination spreadsheet ID | One literal in the configuration | Replaced with private Script Property lookup |
| Personal email address | One literal in the configuration; repeated in test fixtures | Replaced with private configuration; synthetic test address |
| Sample-specific reference | A help message referred to a particular thesis sample size | Removed the reference |
| API keys, passwords, OAuth tokens, private cryptographic keys | No such embedded credentials found by source review and pattern checks | No credential values included |
| Internal keys | State and metadata property names are present | Retained; these are labels, not authentication secrets |
| Participant-level records or actual statistical results | None found embedded in the delivered production source | No participant dataset or backup files included |
| Study/treatment names | Original labels and text aliases identified the studies and interventions | Replaced throughout with generic study/group labels; numeric input codes retained |
| Analysis schema | Clinical variable names, categories, follow-up and analysis settings are present | Retained and disclosed in README |
| Tests | Artificial records used to exercise the code | Converted configuration identifiers to synthetic values |

## Generic labels

All original study and intervention names were removed from the source, including aliases used to recognize textual input, CONSORT captions, table labels, and highlight configuration. Labels now use Estudo A / Estudo B and Grupo 1–4. Clinical variable names and the numeric group structure remain visible.

## Configuration change

The public source reads `RRT_AUTHORIZED_EMAIL` and `RRT_SPREADSHEET_ID` from Apps Script Script Properties. The real values must be supplied privately in the existing Google project. Execution stops if either value is missing.

## Checks performed

- Reviewed the configuration blocks and external-service calls.
- Searched for spreadsheet/Drive links, email addresses, long identifier-like literals, and common credential patterns.
- Searched for sample-specific help text and embedded participant-record arrays.
- Verified syntax and ran 25 simulated regression scenarios.
- Verified that the known private identifiers and removed sample reference are absent from all delivered public text files.

This is a scoped source review, not a guarantee that every detail of the research design is confidential or that a previously published repository has no sensitive history. The retained methodology is intentionally visible.

## Final generic-name pass

Institutional/project references were replaced in code, property names, backup names, comments, regular expressions, tests, documentation, and filenames. Development dates were removed from module-version identifiers. Public minimum-wage reference dates remain for calculation purposes.
