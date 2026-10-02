# Mock data generator review

Source: MatheusMacedoDev/onlyoffice.github.io, commit 3c65d7cdd.
The student's function is included without logic edits in AI 4.0.5.

## Requested changes

- Quoted metadata, display label and all 14 helper translations: integrated.
- Rows schema and runtime agree on 1-100; omission defaults to 10.
- One header-reading editor command; normalization is outside it.
- Trailing empty headers are trimmed, the address is resized, and more than 50 fields are rejected.
- Early target check works for values, whitespace, zero, false and empty-result formulas. An edit during the AI request is caught by the second check.
- Error spacing, indentation, retry hints and range description are updated. The range description says single-row, while implementation deliberately normalizes multi-row selections to the first row.
- A1 describes an address notation, not a requirement that the header starts at cell A1. The offset AA7 test passed.

## Remaining findings

1. Lines 328-347: the final target check and writes use separate callCommand invocations. A change after that check can be overwritten. Keep the early check, but combine the final validation and writes in one editor command, scanning the complete target before any SetValue.
2. Lines 125, 298 and 337: each stage uses GetActiveSheet instead of retaining the source worksheet. Switching sheets during generation can write into the newly selected sheet. Capture and re-resolve the original sheet, and fail if it is unavailable.
3. Lines 268-289: StartAction(GroupActions) is outside the guarded region. If it rejects after Block starts, Block is not ended. Track successfully started actions and clean them up even when startup fails.

## Tests

Run `node test-matheus-mock.cjs` from this directory for source tests.
Run `node test-matheus-mock.cjs ../../scripts/helpers/helpers.js --bundle` for the generated bundle.
Both runs: 33 passed, 3 failed. Failures intentionally remain visible; the student source was not patched to make them pass.
The test harness models separate serialized editor commands and 1-based Row/Col, checked against ONLYOFFICE sdkjs/cell/apiBuilder.js. It simulates provider results, concurrent edits and sheet switches. These are not live-editor tests and do not establish model-routing or retry behavior.
Historical regression scripts in this directory target earlier, hardened store implementations and are not a substitute for this revision-specific suite.

Test this build on disposable worksheets. Do not switch worksheets or edit the target while generation is running. No changes were made to the existing slide-note tool.
