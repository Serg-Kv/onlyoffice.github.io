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
		"name": "addNoteToSlide",
		"text": "Insert Note",
		"description": "Adds speaker notes to a presentation slide. Use for requests to add, write, generate or summarize talking points, presenter notes, or a speaking script. Defaults to the currently open slide when no slide number is given, including requests for this slide. Use text for exact wording or request to generate notes from the slide text and tables. Call once per slide when targeting multiple slides.",
		"parameters": {
			"type": "object",
			"properties": {
				"slideNumber": {
					"type": "integer",
					"description": "Optional one-based slide number. Omit for this slide or the current slide. Defaults to the currently open slide.",
					"minimum": 1
				},
				"text": {
					"type": "string",
					"description": "Exact text to append to speaker notes. Use only when the user supplies the wording; do not also pass request."
				},
				"request": {
					"type": "string",
					"description": "Instructions to generate talking points, speaker notes or a speaking script from the slide content. Do not also pass text."
				}
			},
			"required": []
		},
		"examples": [
			{
				"prompt": "add a note with the following content to slide 3: Hello, world!",
				"arguments": { "slideNumber": 3, "text": "Hello, world!" }
			},
			{
				"prompt": "add talking points to slide 2",
				"arguments": { "slideNumber": 2, "request": "add talking points to slide 2" }
			},
			{ "prompt": "generate talking points for this slide", "arguments": { "request": "Generate talking points from this slide" } },
			{ "prompt": "generate talking points for slide 2", "arguments": { "slideNumber": 2, "request": "Generate talking points from the slide" } },
			{ "prompt": "Write speaker notes for the current slide", "arguments": { "request": "Write speaker notes from this slide" } },
			{ "prompt": "Create a short speaking script for this slide", "arguments": { "request": "Create a short speaking script from this slide" } },
			{ "prompt": "Summarize the table on this slide in the notes", "arguments": { "request": "Summarize the slide table in speaker notes" } },
			{ "prompt": "Add a note to this slide: Remember to thank the audience", "arguments": { "text": "Remember to thank the audience" } },
		]
	});

	func.call = async function (params) {
		console.log("[addNoteToSlide] start", params);
		try {
			params = params || {};
			if (params.slideNumber !== undefined && (!Number.isInteger(params.slideNumber) || params.slideNumber < 1)) {
				throw new window.AgentState.ToolError("slideNumber must be a positive integer.");
			}
			for (const key of ["text", "request"]) {
				if (params[key] !== undefined && (typeof params[key] !== "string" || !params[key].trim())) {
					throw new window.AgentState.ToolError(key + " must be a non-empty string.");
				}
			}
			console.log("[addNoteToSlide] mode", params.request ? "AI-generated notes" : "literal text");
			Asc.scope.params = params;
			// Read, compute and validate parameters
			let callResult = await Asc.Editor.callCommand(function () {
				let presentation = Api.GetPresentation();
				let slide;
				let slideContent;
				if (!Asc.scope.params.text && !Asc.scope.params.request) {
					return { error: "missing_text" };
				}
				if (Asc.scope.params.text && Asc.scope.params.request) {
					return { error: "invalid_text_and_request" };
				}

				if (Asc.scope.params.slideNumber !== undefined) {
					slide = presentation.GetSlideByIndex(Asc.scope.params.slideNumber - 1);
					if (!slide) return { error: "slide_not_found", slidesCount: presentation.GetSlidesCount() };
				}
				else {
					slide = presentation.GetCurrentSlide();
				}

				if (!slide) return { error: "no_current_slide" };
				console.log("[addNoteToSlide] slide resolved", slide.GetSlideIndex());
				if (typeof slide.AddNotesText !== "function") return { error: "notes_api_unavailable" };
				if (Asc.scope.params.text) {
					if (!slide.AddNotesText(Asc.scope.params.text)) return { error: "failed_to_add_note" };
					return { status: "ok", slideIndex: slide.GetSlideIndex(), textLength: Asc.scope.params.text.length };
				}

				// Fetch slide content for LLM case
				if (Asc.scope.params.request) {

					// Get slide content. Tolerate errors.
					let shapesContent = [];
					let shapes = slide.GetAllShapes();
					for (let i = 0; i < shapes.length; i++) {
						let shape = shapes[i];
						let shapeText = "";
						try {
							let content = shape.GetDocContent();
							if (content) {
								let count = content.GetElementsCount();
								let parts = [];
								for (let j = 0; j < count; j++) {
									let el = content.GetElement(j);
									if (el && el.GetText) {
										parts.push(el.GetText());
									}
								}
								shapeText = parts.join("\n");
							}
						}
						// Tolerate failures reading slide content
						catch (e) {
							console.warn("[addNoteToSlide] could not read shape", i, String(e));
						}
						if (shapeText) shapesContent.push(shapeText);
					}

					let shapesResult = shapesContent.join("\n\n");

					// Get slide content from tables. Tolerate errors.
					let tableResults = [];
					try {
						let aTables = slide.GetAllTables();
						for (let i = 0; i < aTables.length; i++) {
							let table = aTables[i];
							let rows = [];
							let k = 0;
							let rowObj = table.GetRow(k++);
							while (rowObj) {
								let row = [];
								for (let c = 0; c < rowObj.GetCellsCount(); c++) {
									let cell = rowObj.GetCell(c);
									let text = "";
									if (cell && typeof cell.GetText === "function") {
										text = cell.GetText();
									} else if (cell && cell.GetContent) {
										let content = cell.GetContent();
										if (content && content.GetText) text = content.GetText();
									}
									row.push(text);
								}
								rows.push(row);
								rowObj = table.GetRow(k++);
							}
							tableResults.push(rows);
						}
					}
					catch (e) {
						console.warn("[addNoteToSlide] could not read tables", String(e));
					}
					let tableJsonContents = JSON.stringify(tableResults);

					slideContent = "Plain text of the slide: " + shapesResult + "\n\n" + "Contents of tables on the slide: " + tableJsonContents;
					console.log("[addNoteToSlide] extracted context", slideContent);
				}
				return {
					slideContentObj: slideContent,
					slideIndex: slide.GetSlideIndex()
				};
			});
			console.log("[addNoteToSlide] slide lookup result", callResult);
			if (callResult && callResult.error === "notes_api_unavailable")
				throw new window.AgentState.ToolError("This editor does not provide slide.AddNotesText.");
			if (callResult && callResult.error === "failed_to_add_note")
				throw new window.AgentState.ToolError("Failed to add the note.");
			if (!callResult || callResult.error === "no_current_slide") {
				throw new window.AgentState.ToolError("No current slide is available.");
			}

			if (callResult && callResult.error === "slide_not_found") {
				throw new window.AgentState.ToolError("Slide " + params.slideNumber + " does not exist! The presentation has " + callResult.slidesCount + " slides.");
			}
			if (callResult && callResult.error === "missing_text") {
				throw new window.AgentState.ToolError("No text was passed to the addNoteToSlide");
			}
			if (callResult && callResult.error === "invalid_text_and_request") {
				throw new window.AgentState.ToolError("Pass either text or request, not both.");
			}


			if (params.text) {
				console.log("[addNoteToSlide] completed in one editor call", callResult);
				return callResult;
			}
			// Resolve once so a selection change during generation cannot redirect the note.
			Asc.scope.noteTargetIndex = callResult.slideIndex;
			Asc.scope.params = { ...params };
			var text = Asc.scope.params.text;

			if (Asc.scope.params.request) {

				// Create LLM request
				let llmPrompt =
					`You are an AI chatbox. You are tasked to generate notes to a specific slide of a presentation.
						To do that, you should primarily follow the user's request which is: ${Asc.scope.params.request}
						To enrich your output, you should use the slide's content: ${callResult.slideContentObj}
						Note that the slide contents and tables, may be empty.
						Do note make stuff up. If there is not enough context to generate notes, simply return "Not enough content"
						If the request and presentation are not in english try to detect the language and match it in your output.
						`
				let requestEngine = AI.Request.create(AI.ActionType.Chat);
				if (!requestEngine)
					throw new window.AgentState.ToolError("No Chat model is configured for generating notes.");
				console.log("[addNoteToSlide] AI request", { model: requestEngine.modelUI.name, prompt: llmPrompt });

				let isSendedEndLongAction = false;
				async function checkEndAction() {
					if (!isSendedEndLongAction) {
						await Asc.Editor.callMethod("EndAction", ["Block", "AI (" + requestEngine.modelUI.name + ")"]);
						isSendedEndLongAction = true;
					}
				}

				await Asc.Editor.callMethod("StartAction", ["Block", "AI (" + requestEngine.modelUI.name + ")"]);
				let groupStarted = false;

				try {
					await Asc.Editor.callMethod("StartAction", ["GroupActions"]);
					groupStarted = true;
					text = await requestEngine.chatRequest(llmPrompt, false);
				} catch (error) {
					throw new window.AgentState.ToolError("AI note generation failed: " + (error && error.message || String(error)));
				} finally {
					try {
						await checkEndAction();
					} finally {
						if (groupStarted) await Asc.Editor.callMethod("EndAction", ["GroupActions"]);
					}
					console.log("[addNoteToSlide] AI actions released");
				}
				console.log("[addNoteToSlide] AI response", text);
			}
			if (typeof text !== "string" || !text.trim()) {
				throw new window.AgentState.ToolError("No note text was produced.");
			}
			console.log("[addNoteToSlide] inserting note", { slideNumber: params.slideNumber, text: text });
			Asc.scope.addNotesResult = text;
			callResult = await Asc.Editor.callCommand(function () {
				// Push result to notes
				let text = Asc.scope.addNotesResult;
				let presentation = Api.GetPresentation();
				let slide = presentation.GetSlideByIndex(Asc.scope.noteTargetIndex);
				if (!slide) return { error: "no_current_slide" };
				if (typeof slide.AddNotesText !== "function") return { error: "notes_api_unavailable" };
				if (!slide.AddNotesText(text)) {
					return { error: "failed_to_add_note", text: text, slideNumber: slide.GetSlideIndex() }
				}
				return { status: "ok", slideIndex: slide.GetSlideIndex(), textLength: text.length };
			});
			console.log("[addNoteToSlide] insertion result", callResult);
			if (!callResult || callResult.error === "no_current_slide") {
				throw new window.AgentState.ToolError("No current slide is available for note insertion.");
			}
			if (callResult.error === "notes_api_unavailable") {
				throw new window.AgentState.ToolError("This editor does not provide slide.AddNotesText.");
			}

			if (callResult && callResult.error === "failed_to_add_note") {
				throw new window.AgentState.ToolError("failed to add note. Parametes: Text: " + callResult.text + ", slideNumber: " + callResult.slideNumber);
			}
			if (callResult && callResult.error === "slide_not_found") {
				throw new window.AgentState.ToolError("Slide " + params.slideNumber + " does not exist! The presentation has " + callResult.slidesCount + " slides.");
			}
			console.log("[addNoteToSlide] completed", callResult);
			return callResult;
		} catch (error) {
			console.error("[addNoteToSlide] failed", error);
			throw error;
		} finally {
			delete Asc.scope.params;
			delete Asc.scope.addNotesResult;
			delete Asc.scope.noteTargetIndex;
		}
	};

	return func;
})();
