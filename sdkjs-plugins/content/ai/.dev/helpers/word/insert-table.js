(function(){
	let func = new RegisteredFunction({
		name: "insertTable",
		description: "Use this function to insert a table at the current cursor position or at the start/end of the document. You can specify the number of rows and columns, and optionally add headers.",
		parameters: {
			type: "object",
			properties: {
				rows: { type: "number", description: "number of rows in the table" },
				columns: { type: "number", description: "number of columns in the table" },
				hasHeaders: { type: "boolean", description: "whether the first row should be formatted as headers" },
				tableStyle: { type: "string", description: "optional table style name (e.g., 'Table Grid', 'Light Grid')" },
				width: { type: "number", description: "table width percentage (default is 100)" },
				widthType: { type: "string", description: "width type - 'percent' or 'point' (default is 'percent')" },
				position: { type: "string", description: "where to insert the table - 'current', 'start', or 'end' (default is 'current')" }
			},
			required: ["rows", "columns"]
		}
	});

	func.call = async function(params) {
		Asc.scope.rows = params.rows || 3;
		Asc.scope.columns = params.columns || 3;
		Asc.scope.hasHeaders = params.hasHeaders || false;
		Asc.scope.tableStyle = params.tableStyle;
		Asc.scope.width = params.width || 100;
		Asc.scope.widthType = params.widthType || "percent";
		Asc.scope.position = params.position || "current";

		await Asc.Editor.callCommand(function() {
			let doc = Api.GetDocument();
			
			if (Asc.scope.position === "start") {
				doc.MoveCursorToStart();
			} else if (Asc.scope.position === "end") {
				doc.MoveCursorToEnd();
			}
			
			let table = Api.CreateTable(Asc.scope.rows, Asc.scope.columns);
			doc.InsertContent([table]);
		
			let unit = (Asc.scope.widthType === "point") ? "twips" : Asc.scope.widthType;
			let widthValue = (Asc.scope.widthType === "point") ? (Asc.scope.width * 20) : Asc.scope.width;
			table.SetWidth(unit, widthValue);
		
			if (Asc.scope.tableStyle) {
				table.SetStyle(Asc.scope.tableStyle);
			}
		
			if (Asc.scope.hasHeaders) {
				for (let col = 0; col < Asc.scope.columns; col++) {
					let cell = table.GetCell(0, col);
					if (cell) {
						let para = cell.GetContent().GetElement(0);
						if (para) {
							let textPr = para.GetTextPr();
							textPr.SetBold(true);
							para.SetTextPr(textPr);
						}
					}
				}
			}
		});
	};

	return func;
})();