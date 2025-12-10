(function(){
	let func = new RegisteredFunction({
		name: "changeTextStyle",
		description: "Use this function to change text formatting including bold, italic, underline, strikeout, font size, and text case.",
		parameters: {
			type: "object",
			properties: {
				bold: { type: "boolean", description: "whether to make the text bold" },
				italic: { type: "boolean", description: "whether to make the text italic" },
				underline: { type: "boolean", description: "whether to underline the text" },
				strikeout: { type: "boolean", description: "whether to strike out the text" },
				fontSize: { type: "number", description: "font size to apply to the selected text" },
				caseType: { type: "string", description: "'upper' for UPPERCASE, 'lower' for lowercase, 'sentence' for Sentence case, 'capitalize' for Capitalize Each Word, 'toggle' for tOGGLE cASE" }
			}
		}
	});

	func.call = async function(params) {
		Asc.scope.bold = params.bold;
		Asc.scope.italic = params.italic;
		Asc.scope.underline = params.underline;
		Asc.scope.strikeout = params.strikeout;
		Asc.scope.fontSize = params.fontSize;
		Asc.scope.caseType = params.caseType;
		await Asc.Editor.callCommand(function(){
			let doc = Api.GetDocument();
			let range = doc.GetRangeBySelect();
			if (!range || "" === range.GetText())
			{
				doc.SelectCurrentWord();
				range = doc.GetRangeBySelect();
			}

			if (!range)
				return;

			if (undefined !== Asc.scope.bold)
				range.SetBold(Asc.scope.bold);

			if (undefined !== Asc.scope.italic)
				range.SetItalic(Asc.scope.italic);

			if (undefined !== Asc.scope.underline)
				range.SetUnderline(Asc.scope.underline);

			if (undefined !== Asc.scope.strikeout)
				range.SetStrikeout(Asc.scope.strikeout);

			if (undefined !== Asc.scope.fontSize)
				range.SetFontSize(Asc.scope.fontSize);

			// Case Type
			if (undefined !== Asc.scope.caseType) {
				let text = range.GetText();
				
				if (!text || text.trim() === "") {
					text = doc.GetCurrentWord();
					if (text) {
						doc.SelectCurrentWord();
						range = doc.GetRangeBySelect();
					}
				}

				if (text && text.trim() !== "") {
					// Define case conversion functions
					let convertCase;
					switch (Asc.scope.caseType) {
						case "upper":
							convertCase = t => t.toUpperCase();
							break;
						case "lower":
							convertCase = t => t.toLowerCase();
							break;
						case "sentence":
							convertCase = t => t.charAt(0).toUpperCase() + t.slice(1).toLowerCase();
							break;
						case "capitalize":
							convertCase = t => t.replace(/\b\w/g, l => l.toUpperCase());
							break;
						case "toggle":
							convertCase = t => t.split('').map(c => c === c.toUpperCase() ? c.toLowerCase() : c.toUpperCase()).join('');
							break;
						default:
							convertCase = t => t;
					}

					// Paragraph processing function
					const processParagraphs = paragraphs => {
						for (let i = 0; i < paragraphs.length; i++) {
							const para = paragraphs[i];
							
							if (!para.GetElementsCount) continue;

							const elementsCount = para.GetElementsCount();
							let fullText = "";
							let runs = [];

							for (let j = 0; j < elementsCount; j++) {
								const elem = para.GetElement(j);
								if (elem.GetText) {
									const text = elem.GetText();
									if (text) {
										fullText += text;
										runs.push({ element: elem, text: text, length: text.length });
									}
								}
							}
							
							if (fullText.trim() === "") continue;

							const newFullText = convertCase(fullText);

							if (newFullText !== fullText) {
								para.RemoveAllElements();
								let currentPos = 0;
								for(let k = 0; k < runs.length; k++) {
									const run = runs[k];
									const newRunText = newFullText.substring(currentPos, currentPos + run.length);
									const newRun = Api.CreateRun();
									
									const oldPr = run.element.GetTextPr();
									newRun.SetTextPr(oldPr);
									newRun.AddText(newRunText);
									
									para.AddElement(newRun);
									currentPos += run.length;
								}
							}
						}
					};

					if (range && range.GetText && range.GetText().trim() !== "") {
						processParagraphs(range.GetAllParagraphs());
					} else {
						processParagraphs(doc.GetAllParagraphs());
					}
				}
			}

		});
	};

	return func;
})();