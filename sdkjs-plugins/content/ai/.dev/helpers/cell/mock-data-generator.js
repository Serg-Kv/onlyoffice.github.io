/*
 * (c) Copyright Ascensio System SIA 2010-2025
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA at 20A-6 Ernesta Birznieka-Upish
 * street, Riga, Latvia, EU, LV-1050.
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * Pursuant to Section 7(b) of the License you must retain the original Product
 * logo when distributing the program. Pursuant to Section 7(e) we decline to
 * grant you any rights under trademark law for use of our trademarks.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

(function () {
    let func = new RegisteredFunction({
        "name": "mockDataGenerator",
        "text": "Generate mock data",
        description:
            "Generate mock data for a selected table header with type infer based on the name of each field. If no header is selected, works with header position provided by prompt parameter.",
        parameters: {
            type: "object",
            properties: {
                range: {
                    type: "string",
                    description: "Cell range with the table header (e.g., 'A1:C1'). If omitted, uses the selected header.",
                },
                rows: {
                    type: "number",
                    description: "Amount of rows to fill with generated mock data.",
                    default: 10,
                },
            },
            required: [],
        },
        examples: [
            {
                prompt: "Generate data for the selected table header",
                arguments: {},
            },
            {
                prompt: "Fill the table below the current header with realistic fake data",
                arguments: {},
            },
            {
                prompt: "Generate 20 rows of mock data for the selected header",
                arguments: { rows: 20 },
            },
            {
                prompt: "Create sample data with 5 rows for the header range A1:C1",
                arguments: { range: "A1:C1", rows: 5 },
            }
        ]
    });

    const getHeaderFromSelection = async function () {
        const header = await Asc.Editor.callCommand(function () {
            const worksheet = Api.GetActiveSheet();
            const selection = worksheet.Selection;

            if (!selection) {
                console.log("[mockDataGenerator] getHeaderFromSelection: no active selection");
                return null;
            }

            const result = {
                address: selection.GetAddress(true, true, "xlA1"),
                fields: (selection.GetValue2() || [])[0] || []
            }

            console.log("[mockDataGenerator] getHeaderFromSelection (inside callCommand):", result);

            return result;
        })

        console.log("[mockDataGenerator] getHeaderFromSelection result:", header);

        return header;
    }

    const getHeaderFromRangeProperty = async function (range) {
        if (typeof range !== "string" || !range.trim()) {
            console.log("[mockDataGenerator] getHeaderFromRangeProperty: no range provided, skipping");
            return null;
        }

        Asc.scope.range = range.trim();

        const header = await Asc.Editor.callCommand(function () {
            const worksheet = Api.GetActiveSheet();
            const headerRange = worksheet.GetRange(Asc.scope.range);

            if (!headerRange) {
                console.log("[mockDataGenerator] getHeaderFromRangeProperty: range not found:", Asc.scope.range);
                return null;
            }

            const result = {
                address: headerRange.GetAddress(true, true, "xlA1"),
                fields: (headerRange.GetValue2() || [])[0] || []
            }

            console.log("[mockDataGenerator] getHeaderFromRangeProperty (inside callCommand):", result);

            return result;
        })

        console.log("[mockDataGenerator] getHeaderFromRangeProperty result:", header);

        return header;
    }

    const getHeader = async function (range) {
        console.log("[mockDataGenerator] getHeader: resolving header, range param =", range);

        let header = await getHeaderFromRangeProperty(range);

        if (!header)
            header = await getHeaderFromSelection();

        console.log("[mockDataGenerator] getHeader: final resolved header =", header);

        return header;
    }

    const parseMatrixFromAIResponse = function (aiResponse, rowsAmount, columnsAmount) {
        console.log("[mockDataGenerator] parseMatrixFromAIResponse: raw AI response =", aiResponse);
        console.log("[mockDataGenerator] parseMatrixFromAIResponse: expected shape =", rowsAmount, "rows x", columnsAmount, "columns");

        if (!aiResponse) {
            console.log("[mockDataGenerator] parseMatrixFromAIResponse: empty AI response, returning null");
            return null;
        }

        const matchedArrays = aiResponse.match(/\[[\s\S]*\]/);

        console.log("[mockDataGenerator] parseMatrixFromAIResponse: regex match result =", matchedArrays);

        try {
            const parsed = JSON.parse(matchedArrays[0]);

            console.log("[mockDataGenerator] parseMatrixFromAIResponse: parsed JSON =", parsed);

            if (!Array.isArray(parsed)
                || !parsed.every(row => Array.isArray(row))
                || parsed.length !== rowsAmount
                || parsed.some(row => row.length !== columnsAmount)
            ) {
                console.log("[mockDataGenerator] parseMatrixFromAIResponse: parsed matrix has an unexpected shape");
                return null;
            }

            return parsed;

        } catch (error) {
            console.log("[mockDataGenerator] parseMatrixFromAIResponse: failed to parse AI response as JSON", error);
            throw new window.AgentState.ToolError("Failed to parse AI response as JSON.");
        }
    }

    const generateMockMatrix = async function (fields, rows) {
        console.log("[mockDataGenerator] generateMockMatrix: fields =", fields, "rows =", rows);

        const mappedFields = fields.map(field => field === "" ? "[Empty]" : field);

        const argPrompt = [
            "You are a mock data generator for a spreadsheet table.",
            `Column names in order are: ${mappedFields.join(", ")}.`,
            `Generate ${rows} rows of realistic mock data for each column, based on the column name.` +
            "Each value must match the meaning of its column (e.g. if the column is 'Email', generate realistic email addresses).",
            "If a column is marked as '[Empty]', consider that all column values should be empty strings.",
            "Strict rules:",
            `1. Return only a JSON array of arrays (row-major), containing exactly ${rows} rows x ${fields.length} columns.`,
            "2. No markdown, no code fences, no explanations, no extra text, only the JSON array.",
            '3. Format example: [["cell_1_1", "cell_1_2"], ["cell_2_1", "cell_2_2"]]',
        ].join("\n");

        console.log("[mockDataGenerator] generateMockMatrix: prompt sent to AI =\n" + argPrompt);

        const requestEngine = AI.Request.create(AI.ActionType.Chat);

        if (!requestEngine) {
            console.log("[mockDataGenerator] generateMockMatrix: AI.Request.create returned no engine");
            throw new window.AgentState.ToolError("AI Request engine is not available.");
        }

        console.log("[mockDataGenerator] generateMockMatrix: using model =", requestEngine.modelUI.name);

        let isSendedEndLongAction = false;

        async function checkEndAction() {
            if (!isSendedEndLongAction) {
                await Asc.Editor.callMethod("EndAction", [
                    "Block",
                    `AI (${requestEngine.modelUI.name})`
                ]);

                isSendedEndLongAction = true;

                console.log("[mockDataGenerator] generateMockMatrix: EndAction (Block) sent");
            }
        }

        await Asc.Editor.callMethod("StartAction", [
            "Block",
            `AI (${requestEngine.modelUI.name})`
        ])
        await Asc.Editor.callMethod("StartAction", ["GroupActions"]);

        console.log("[mockDataGenerator] generateMockMatrix: StartAction sent, sending chatRequest...");

        let aiResult;
        try {
            aiResult = await requestEngine.chatRequest(argPrompt, false);
        } finally {
            await checkEndAction();
            await Asc.Editor.callMethod("EndAction", ["GroupActions"]);
        }

        console.log("[mockDataGenerator] generateMockMatrix: chatRequest resolved, raw result =", aiResult);


        const matrix = parseMatrixFromAIResponse(aiResult, rows, fields.length);

        console.log("[mockDataGenerator] generateMockMatrix: final matrix =", matrix);

        return matrix;
    }

    const insertMatrixBelowHeader = async function (header, matrix) {
        Asc.scope.address = header.address;
        Asc.scope.matrix = matrix;
        Asc.scope.colCount = (matrix[0] || []).length;
        Asc.scope.rowCount = matrix.length;

        console.log(
            "[mockDataGenerator] insertMatrixBelowHeader: header address =", header.address,
            "rows =", Asc.scope.rowCount, "cols =", Asc.scope.colCount
        );

        await Asc.Editor.callCommand(function () {
            const worksheet = Api.GetActiveSheet();
            const headerRange = worksheet.GetRange(Asc.scope.address);
            const fillRange = headerRange.Resize(Asc.scope.rowCount + 1, Asc.scope.colCount);

            console.log("[mockDataGenerator] insertMatrixBelowHeader (inside callCommand): fillRange address =", fillRange.GetAddress(true, true, "xlA1"));

            for (let rowIndex = 2; rowIndex <= Asc.scope.rowCount + 1; rowIndex++) {
                let row = fillRange.GetRows(rowIndex);
                for (let columnIndex = 1; columnIndex <= Asc.scope.colCount; columnIndex++) {
                    row.GetCells(columnIndex).SetValue(Asc.scope.matrix[rowIndex - 2][columnIndex - 1]);
                }
            }

            console.log("[mockDataGenerator] insertMatrixBelowHeader (inside callCommand): finished filling", Asc.scope.rowCount, "rows");
        })

        console.log("[mockDataGenerator] insertMatrixBelowHeader: callCommand completed");
    }

    func.call = async function (params) {
        console.log("[mockDataGenerator] func.call: invoked with params =", params);

        params = params || {};
        const rows = params.rows === undefined ? 10 : params.rows;
        if (!Number.isInteger(rows) || rows < 1 || rows > 1000) {
            throw new window.AgentState.ToolError("Row count must be an integer between 1 and 1000.");
        }
        const header = await getHeader(params.range);

        if (!header || header.fields.length === 0) {
            console.log("[mockDataGenerator] func.call: no header resolved, aborting", header);
            throw new window.AgentState.ToolError("No header selected or found in the current worksheet.");
        }

        const fields = header.fields
            .map(field =>
                String(field === null || field === undefined ? "" : field).trim()
            );

        console.log("[mockDataGenerator] func.call: normalized fields =", fields);

        const nonEmptyFields = fields.filter(field => field !== "");

        if (nonEmptyFields.length === 0) {
            console.log("[mockDataGenerator] func.call: header contains only empty fields, aborting");
            throw new window.AgentState.ToolError("The selected header contains only empty fields.");
        }

        const matrix = await generateMockMatrix(fields, rows);

        if (!matrix) {
            console.log("[mockDataGenerator] func.call: generateMockMatrix returned no matrix, aborting");
            throw new window.AgentState.ToolError("AI returned an invalid matrix shape");
        }

        await insertMatrixBelowHeader(header, matrix);

        const result = {
            status: "ok",
            headers: header.fields,
            columns: header.fields.length,
            generatedRows: rows,
        }

        console.log("[mockDataGenerator] func.call: completed, result =", result);

        return result;
    };

    return func;
})();
