(function(){
	let func = new RegisteredFunction({
		name: "insertList",
		description: "Use this function to create simple numbered or bulleted lists at the current cursor position or at the start/end of the document.",
		parameters: {
			type: "object",
			properties: {
				items: { type: "array", description: "array of strings representing list items", items: { type: "string" } },
				listType: { type: "string", description: "'numbered' for numbered list, 'bulleted' for bulleted list (default is 'bulleted')" },
				position: { type: "string", description: "where to insert the list - 'current', 'start', or 'end' (default is 'current')" }
			},
			required: ["items"]
		}
	});

	func.call = async function(params) {
		Asc.scope.items = params.items || ["Item 1", "Item 2", "Item 3"];
		Asc.scope.listType = params.listType || "bulleted";
		Asc.scope.position = params.position || "current";

		await Asc.Editor.callCommand(function() {
			let doc = Api.GetDocument();
			
			if (Asc.scope.position === "start") {
				doc.MoveCursorToStart();
			} else if (Asc.scope.position === "end") {
				doc.MoveCursorToEnd();
				let newParagraph = Api.CreateParagraph();
				doc.InsertContent([newParagraph]);
			} else if (Asc.scope.position === "current") {
				let newParagraph = Api.CreateParagraph();
				doc.InsertContent([newParagraph]);
			}
			
			let paragraphs = [];
			let numbering;
			
			if (Asc.scope.listType === "numbered") {
				numbering = doc.CreateNumbering("numbered");
			} else {
				numbering = doc.CreateNumbering("bullet");
			}
			
			let numLvl = numbering.GetLevel(0);
			
			for (let i = 0; i < Asc.scope.items.length; i++) {
				let item = Asc.scope.items[i];
				let paragraph = Api.CreateParagraph();
				paragraph.AddText(item);
				paragraph.SetNumbering(numLvl);
				paragraph.SetContextualSpacing(true);
				paragraphs.push(paragraph);
			}
			
			doc.InsertContent(paragraphs);
		});
	};

	return func;
})();