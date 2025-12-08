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
    name: "describeImage",
    description:
      "Allows users to select an image and generate a meaningful title, description, caption, or alt text for it using AI.",
    // Define parameters so the AI knows what to ask for
    parameters: {
      type: "object",
      properties: {
        prompt: {
          type: "string",
          description:
            "instruction for the AI (e.g., 'Add a short title for this chart.')",
        },
      },
      required: ["prompt"],
    },
    // Provide examples to train the AI on usage
    examples: [
      {
        prompt: "Add a short title for this chart.",
        arguments: { prompt: "Add a short title for this chart." },
      },
      {
        prompt: "Write me a 1–2 sentence description of this photo.",
        arguments: {
          prompt: "Write me a 1–2 sentence description of this photo.",
        },
      },
      {
        prompt: "Generate a descriptive caption for this organizational chart.",
        arguments: {
          prompt:
            "Generate a descriptive caption for this organizational chart.",
        },
      },
      {
        prompt: "Provide accessibility-friendly alt text for this infographic.",
        arguments: {
          prompt:
            "Provide accessibility-friendly alt text for this infographic.",
        },
      },
    ],
  });

   // The actual logic executed inside the editor
  func.call = async function (params) {
    let prompt = params.prompt;

    async function insertMessage(message) {
      Asc.scope._message = String(message || "");
      // 3. Insert the result into the document
      await Asc.Editor.callCommand(function () {
        const msg = Asc.scope._message || "";
        const doc = Api.GetDocument();
        const selected =
          (doc.GetSelectedDrawings && doc.GetSelectedDrawings()) || [];
        if (selected.length > 0) {
          for (let i = 0; i < selected.length; i++) {
            const drawing = selected[i];
            const para = Api.CreateParagraph();
            para.AddText(msg);
            drawing.InsertParagraph(para, "after", true);
          }
        } else {
          const para = Api.CreateParagraph();
          para.AddText(msg);
          let range = doc.GetCurrentParagraph();
          range.InsertParagraph(para, "after", true);
        }
        Asc.scope._message = "";
      }, true);
    }

    try {
    // 1. Get the selected image
      let imageData = await new Promise((resolve) => {
        window.Asc.plugin.executeMethod(
          "GetImageDataFromSelection",
          [],
          function (result) {
            resolve(result);
          }
        );
      });

      if (!imageData || !imageData.src) {
        await insertMessage("Please select a valid image first.");
        return;
      }

      const whiteRectangleBase64 =
        "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==";
      if (imageData.src === whiteRectangleBase64) {
        await insertMessage("Please select a valid image first.");
        return;
      }

      let argPrompt = prompt + " (for the selected image)";
      // 2. Send image + prompt to the AI engine
      let requestEngine = AI.Request.create(AI.ActionType.Chat);
      if (!requestEngine) {
        await insertMessage("AI request engine not available.");
        return;
      }

      const allowVision = /(vision|gemini|gpt-4o|gpt-4v|gpt-4-vision)/i;
      if (!allowVision.test(requestEngine.modelUI.name)) {
        await insertMessage(
          "⚠ This model may not support images. Please choose a vision-capable model (e.g. GPT-4V, Gemini, etc.)."
        );
        return;
      }

      await Asc.Editor.callMethod("StartAction", [
        "Block",
        "AI (" + requestEngine.modelUI.name + ")",
      ]);
      await Asc.Editor.callMethod("StartAction", ["GroupActions"]);

      let messages = [
        {
          role: "user",
          content: [
            { type: "text", text: argPrompt },
            {
              type: "image_url",
              image_url: { url: imageData.src, detail: "high" },
            },
          ],
        },
      ];

      let resultText = "";
      await requestEngine.chatRequest(messages, false, async function (data) {
        if (data) {
          resultText += data;
        }
      });

      await Asc.Editor.callMethod("EndAction", ["GroupActions"]);
      await Asc.Editor.callMethod("EndAction", [
        "Block",
        "AI (" + requestEngine.modelUI.name + ")",
      ]);

      Asc.scope.text = resultText || "";

      if (!Asc.scope.text.trim()) {
        await insertMessage(
          "⚠ AI request failed (maybe the model cannot handle images)."
        );
        return;
      }
      await insertMessage(Asc.scope.text);
    } catch (e) {
      try {
        await Asc.Editor.callMethod("EndAction", ["GroupActions"]);
        await Asc.Editor.callMethod("EndAction", [
          "Block",
          "AI (describeImage)",
        ]);
      } catch (ee) {
        /* ignore */
      }
      console.error("describeImage error:", e);
      await insertMessage(
        "An unexpected error occurred while describing the image."
      );
    }
  };

  return func;
})();


