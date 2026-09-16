# How the Office JavaScript API documentation is generated

The Office JavaScript API reference documentation is generated from TypeScript definition files, code snippets, and repository configuration. The generation pipeline combines standard Rush Stack tools with repository-specific scripts that split and version the definitions, enrich the generated YAML, assemble the published table of contents, and validate the result.

The generated reference files are written to `docs/docs-ref-autogen/`. Do not edit files in that folder directly because the next generation run will overwrite them.

## Content sources

### Type definition files

The API definitions and their TSDoc comments come primarily from these packages in [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped):

- [`office-js/index.d.ts`](https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/office-js/index.d.ts): Release definitions for the Common API, Excel, OneNote, Outlook, PowerPoint, Visio, and Word.
- [`office-js-preview/index.d.ts`](https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/office-js-preview/index.d.ts): Preview definitions for the Common API, Excel, Outlook, PowerPoint, and Word.
- [`custom-functions-runtime/index.d.ts`](https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/custom-functions-runtime/index.d.ts): Excel Custom Functions runtime definitions.
- [`office-runtime/index.d.ts`](https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/office-runtime/index.d.ts): Office Runtime definitions.

The preprocessor supports four source choices:

| Choice | Behavior |
|---|---|
| `DT` | Downloads the DefinitelyTyped files and preserves unchanged API Extractor JSON and API Documenter YAML when possible. |
| `DT+` | Downloads the DefinitelyTyped files and forces a full rebuild. This is the mode used by the scheduled GitHub Action. |
| `CDN` | Downloads the Office.js release and preview definitions from the Office CDN. The Custom Functions and Office Runtime definitions still come from DefinitelyTyped. |
| `Local` | Reads the definition files in `generate-docs/script-inputs/`. Use this mode to test definition changes before submitting them to DefinitelyTyped. |

To test local definitions, copy the modified files to `generate-docs/script-inputs/` using the names expected by the preprocessor, then run the following command from `generate-docs/`.

```bash
./GenerateDocs.sh -b Local
```

You can also run `./GenerateDocs.sh` without `-b` and select **Local files** at the prompt.

### Version-specific definitions

Release documentation is generated for individual API requirement sets. These version-specific definitions are not maintained as independent source files. During every generation run, the [`version-remover`](https://www.npmjs.com/package/versioned-d.ts-tools) command from the `versioned-d.ts-tools` package successively removes APIs associated with newer requirement sets.

`GenerateDocs.sh` defines the version-removal chains for Excel, Outlook, PowerPoint, and Word. The JSON files in `generate-docs/configs/` configure the transformations for each release, online, desktop, and hidden-document variant.

The pipeline also runs the `whats-new` command from the same package. It compares adjacent definition versions and generates the API tables in `docs/includes/` that are included by the requirement-set documentation.

### Code snippets

Code snippets come from two sources:

- The generated Script Lab snippet collection in [OfficeDev/office-js-snippets](https://github.com/OfficeDev/office-js-snippets/blob/prod/snippet-extractor-output/snippets.yaml).
- Local host-specific YAML files in [`docs/code-snippets/`](https://github.com/OfficeDev/office-js-docs-reference/tree/main/docs/code-snippets).

The midprocessor downloads the Script Lab YAML, combines all local snippet files, and merges snippets that target the same API member. It then creates host-specific and version-specific `snippets.yaml` files under `generate-docs/json/`.

Snippet keys use API member UIDs, for example:

```yaml
Excel.Range#values:member:
  - |-
    await Excel.run(async (context) => {
        // ...
    });
```

The Office YAML processor inserts each snippet into the matching generated API item. Snippets are currently emitted in `TypeScript` code fences; the language is not inferred from the snippet contents.

## Running the generation pipeline

Run all commands in this section from `generate-docs/`.

```bash
# Interactive source selection
./GenerateDocs.sh

# Optimized rebuild from DefinitelyTyped
./GenerateDocs.sh -b DT

# Full rebuild from DefinitelyTyped
./GenerateDocs.sh -b DT+

# Build from generate-docs/script-inputs/
./GenerateDocs.sh -b Local
```

`GenerateDocs.sh` installs the root and script dependencies, compiles the TypeScript scripts, records output in `build-log.txt` and errors in `build-errors.txt`, and orchestrates the following stages.

## Generation stages

### 1. Preprocess the definitions

`generate-docs/scripts/preprocessor.ts`:

- Downloads or reads the four source definition files.
- Extracts the Common API and host-specific sections into the corresponding `api-extractor-inputs-*` folders.
- Creates separate preview and release inputs.
- Makes declarations exportable for API Extractor.
- Adds imports needed for Common API, Outlook, and OfficeExtension cross-references.
- Applies targeted fixes needed by the downstream tools.
- Removes affected JSON and YAML output when an input changed, or removes all applicable output during a forced rebuild.

### 2. Generate requirement-set definitions and tables

`GenerateDocs.sh` runs two commands from `versioned-d.ts-tools`:

- `version-remover` creates the definition files for each supported requirement set and special platform variant.
- `whats-new` compares adjacent definition files and writes generated requirement-set tables to `docs/includes/`.

Adding a new requirement set requires updating this orchestration and its related API Extractor configuration, processor version constants, and publishing configuration.

### 3. Run API Extractor

[`@microsoft/api-extractor`](https://api-extractor.com/) reads each prepared `.d.ts` input and writes an API model JSON file under `generate-docs/json/`.

The script skips a host or version when its JSON output folder already exists. The preprocessor and midprocessor remove output folders when changed definitions or snippets require that output to be regenerated.

### 4. Prepare JSON and snippets

`generate-docs/scripts/midprocessor.ts`:

- Repairs canonical references between the Common API, Outlook, OfficeExtension, and the host APIs.
- Cleans enum-member documentation that API Documenter cannot render correctly.
- Downloads and combines Script Lab and local snippets.
- Assigns snippets to the correct host and copies them into every applicable version.
- Copies the Custom Functions API model into the supported Excel outputs.
- Cleans generated Outlook requirement-set include files.

### 5. Run API Documenter

[`@microsoft/api-documenter`](https://api-extractor.com/pages/setup/generating_docs/) converts each API model JSON folder into DocFX YAML under `generate-docs/yaml/`.

The repository uses the standard API Documenter YAML command. Office-specific behavior is applied by the scripts in the following stages rather than by a custom API Documenter extension.

### 6. Apply Office-specific YAML enhancements

`generate-docs/scripts/yaml-office-processor.ts` updates the YAML generated by API Documenter. It:

- Inserts code snippets into matching API members.
- Converts API requirement-set annotations into links to the applicable requirement-set documentation.
- Builds a reverse index from the API Extractor JSON and adds **Used by** sections to referenced types.
- Reports snippets that do not match an API member in the main preview outputs.

### 7. Generate the Outlook item object model tables

`generate-docs/scripts/generate-item-object-model.ts` reads the preview and versioned Outlook API model JSON and generates these include files:

- `docs/includes/outlook-item-object-model-properties.md`
- `docs/includes/outlook-item-object-model-methods.md`
- `docs/includes/outlook-item-object-model-events.md`

These files provide the tables used by the Outlook item object model conceptual page.

### 8. Assemble the publishing output

`generate-docs/scripts/postprocessor.ts`:

- Removes the previous generated reference output, except for retained overview and image content.
- Copies the generated YAML into `docs/docs-ref-autogen/`.
- Combines the generated API Documenter TOCs with the repository's global TOC template.
- Creates TOCs for preview, release, requirement-set, online, desktop, and hidden-document variants.
- Reorganizes enums, OfficeExtension APIs, Office Runtime APIs, Custom Functions APIs, and other special categories.
- Repairs namespace names and links in the generated output.
- Adds links for types contained in type aliases.
- Normalizes generated YAML that would otherwise be formatted incorrectly.

The Open Publishing System uses the files in `docs/docs-ref-autogen/`, `docs/docfx.json`, and `.openpublishing.publish.config.json` to publish the YAML as Microsoft Learn reference pages with the appropriate API-set monikers.

### 9. Update requirement-set page dates

`generate-docs/scripts/update-requirement-set-dates.ts` hashes the generated requirement-set include files and compares them with `generate-docs/script-inputs/include-hashes.json`. When an include changes, the script updates `ms.date` on the requirement-set pages that use that include.

### 10. Validate reference coverage

The final pipeline command runs [`reference-coverage-tester`](https://www.npmjs.com/package/reference-coverage-tester) with `generate-docs/configs/reference-coverage-tester.json`. It checks the generated reference output for missing links or incomplete reference coverage.

## Incremental and full builds

The generated JSON and YAML folders are also the pipeline's incremental-build markers:

- `DT` sets `forceRebuild` to false. If a preprocessed definition is unchanged, its existing output can be reused.
- `DT+`, `CDN`, and `Local` force the preprocessor to invalidate the applicable output.
- A changed host snippet file causes the midprocessor to remove the corresponding YAML output so API Documenter runs again.
- API Extractor and API Documenter skip an output folder when it already exists.

Use `DT+` when validating changes to the pipeline itself or when a complete rebuild is required.

## Automation and preview builds

`.github/workflows/autogen-docs.yml` runs `./GenerateDocs.sh -b DT+` on Tuesdays and Thursdays and when manually dispatched. If generation changes files, the workflow replaces the remote `autogen-docs` branch with the newly generated output for review.

The Open Publishing configuration enables preview builds for pushed branches. Those builds are rendered on `review.learn.microsoft.com`, which is available only to internal Microsoft personnel.
