# Research Results Toolkit

Google Apps Script toolkit for generating dental research tables and CONSORT flow diagrams from Google Sheets, with Kaplan–Meier estimates, log-rank comparisons, and integrated result masking.

## Scope

`RESEARCH_RESULTS.gs` combines five modules into one JavaScript V8 file:

- Sample characteristics, counts, percentages, chi-square and Fisher tests (Table 3).
- Success and survival estimates using Kaplan–Meier, log-log 95% confidence intervals, and global log-rank comparisons (Table 4).
- Individual failure descriptions (Table 5).
- CONSORT flow diagrams and study cutoff controls.
- Masking controls, private backups, and interrupted-write recovery.

This is the results component of an existing spreadsheet-bound research workflow. It does not include the complete master application or a research dataset. Worksheet names follow that project's schema. Public labels use **Estudo A / Estudo B** and **Grupo 1–4**, with the original numeric input codes preserved.

## Public naming

The source, function names, property names, generated backup prefixes, documentation, and file names use the generic toolkit name and `RRT` prefix. Module version strings contain no development dates. Study labels use Estudo A / Estudo B and Grupo 1–4.

The generic namespace is intended for a separately configured installation. It does not read state or backup references from a differently named installation. Do not overwrite a live installation with active masking solely to publish this repository.

Public historical dates used by the minimum-wage conversion table remain because they are calculation reference data, not study dates.

## Private configuration

The public source contains no live spreadsheet ID or personal email address. Set these values privately in **Apps Script → Project Settings → Script Properties**:

| Property | Value to supply privately |
| --- | --- |
| `RRT_AUTHORIZED_EMAIL` | Email address of the authorized account |
| `RRT_SPREADSHEET_ID` | ID of the destination master spreadsheet |

Do not commit the values to this repository. Missing configuration blocks execution. Google authorization and spreadsheet permissions are still required; these property names and the internal state keys are not API credentials.

## Installation

1. Open the master spreadsheet and choose **Extensions → Apps Script**.
2. Copy the entire `RESEARCH_RESULTS.gs` file into a script file with that name.
3. Remove the five older modules replaced by this combined file: `RESULT_MASKING.gs`, `CONSORT.gs`, `TABELA1.gs`, `TABELA2.gs`, and `TABELA3.gs`. Keep one definition of each function.
4. Remove any older duplicate definition of `tabelasAutomatico()`.
5. Preserve the other master-project files and its existing `appsscript.json` manifest.
6. Set the two private Script Properties above.
7. Run `RRT_INSTALAR_E_LIGAR` using the authorized account and grant the required Google permissions.
8. Run `RRT_DIAGNOSTICO`. If writes are pending, run `RRT_REPROCESSAR`.

Retire the previous standalone masking project and any of its triggers. Do not use its restore command after migration. Automated masked generation must run under the authorized account.

## Commands

| Function | Purpose |
| --- | --- |
| `RRT_INSTALAR_E_LIGAR` | Initialize the integration and enable masking. |
| `RRT_LIGAR` | Enable masking and update outputs when required. |
| `RRT_DESLIGAR` | Disable masking and regenerate original results from the current source and cutoff. |
| `RRT_REPROCESSAR` | Rebuild all four outputs while preserving the selected mode. |
| `RRT_DIAGNOSTICO` | Report the selected mode and pending writes. |
| `tabelasAutomatico` | Integration entry point already called by master-project routines. |

Turning masking off regenerates current results; it does not restore an outdated snapshot or manual edits to generated outputs. Diagnostics inspect recorded state, not subsequent manual cell changes.

## Generic study and group labels

| Numeric input code | Public study label |
| --- | --- |
| `0` | Estudo A |
| `1` | Estudo B |

| Numeric input code | Public group label |
| --- | --- |
| `0` | Grupo 1 |
| `1` | Grupo 2 |
| `2` | Grupo 3 |
| `3` | Grupo 4 |

The numeric codes and group membership remain compatible with the existing input schema. Text study aliases accept `Estudo A`, `Estudo B`, `Study A`, `Study B`, `A`, and `B`; the previous identifying aliases are no longer included. Display labels are generic, but clinical field names and analytic methods remain visible.

## Existing generator names

Names are preserved for compatibility with menus and triggers:

| Existing function | Output worksheet |
| --- | --- |
| `GERAR_TABELA1` | TABELA 3 |
| `GERAR_TABELA2` | TABELA 4 |
| `GERAR_TABELA3` | TABELA 5 |
| `GERAR_CONSORT` | CONSORT |

## Masking behavior

- Table 3: replaces selected chi-square/Fisher p-values with five-digit codes and redistributes counts; percentages remain unchanged.
- Table 4: redistributes counts, including counts inside count/percentage text. Log-rank p-values, percentages, times, Kaplan–Meier estimates, and confidence intervals remain unchanged.
- Table 5: independently shuffles each populated field within its study block, including tooth, sex, group, remaining surfaces, vitality, times, failure reason, repair, and current status. Blank cells remain blank.
- CONSORT: redistributes textual `n =` counts while preserving editable panel cells.

Repeated values or fields with only one record can remain unchanged. Independently shuffled fields can produce clinically inconsistent combinations. Masked output must not be used for scientific analysis. This mechanism does not anonymize the source data or remove spreadsheet version history.

The output matrix and rich text are masked before writing. Each masked generation validates a private backup before publishing the selected fields. Backups are created at runtime in the authorized account and are not included in this repository. Pending writes can be rebuilt from the current source.

Google Sheets does not provide an atomic transaction across multiple worksheets. An interruption can leave partial outputs; check diagnostics before using them. Existing backups are not automatically deleted. Project editors can change the source and its authorization checks.

## Statistical methods

Table 4 implements Kaplan–Meier estimation, Greenwood-based log-log 95% confidence intervals, and a global log-rank test between groups within each study, separately for success and survival. It also reports descriptive statistics, counts, and percentages. Cox regression is not implemented.

For comparisons with R, explicitly align the confidence-interval transformation, time units, event coding, study groups, exclusions, and cutoff rules.

## Data disclosure

Study and treatment names have been replaced with generic labels throughout the public source, including text aliases, CONSORT titles, table headers, and highlighting configuration. The source retains clinical variable names, category definitions, methodological settings, and the numeric group structure. These describe the analysis schema; they are not participant-level observations or reported study results. Public reference constants, such as historical minimum-wage values used by the income calculation, also remain in the code.

The previous embedded spreadsheet identifier, personal email address, sample-specific help-text reference, and identifying study/treatment labels were removed. All test records and configuration values are synthetic.

## Local validation

```sh
node tests/research_results_regression.cjs
```

The combined source passes syntax validation and 25 regression scenarios with simulated Google services, including missing-configuration blocking and Table 5 masking. These checks do not constitute execution in the live spreadsheet or independent statistical validation.
