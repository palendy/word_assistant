export async function readDocumentStructure(): Promise<string> {
  return Word.run(async (context) => {
    const body = context.document.body;

    // Load paragraphs with style info
    const paragraphs = body.paragraphs;
    paragraphs.load("items");
    await context.sync();

    for (const para of paragraphs.items) {
      para.load("text");
      para.load("style");
      para.load("alignment");
    }

    const tables = body.tables;
    tables.load("items");
    await context.sync();

    const parts: string[] = [];
    parts.push("=== DOCUMENT STRUCTURE ===");
    parts.push(`Total paragraphs: ${paragraphs.items.length}`);
    parts.push(`Total tables: ${tables.items.length}`);
    parts.push("");

    // Paragraph details with styles
    parts.push("=== PARAGRAPHS ===");
    for (let i = 0; i < paragraphs.items.length; i++) {
      const p = paragraphs.items[i];
      const text = p.text.trim();
      if (!text) {
        parts.push(`[P${i}] (empty) | style: ${p.style}`);
      } else {
        parts.push(`[P${i}] "${text}" | style: ${p.style} | align: ${p.alignment}`);
      }
    }

    // Table details with headers and cell values
    if (tables.items.length > 0) {
      for (let i = 0; i < tables.items.length; i++) {
        const table = tables.items[i];
        table.load("rowCount");
        table.load("values");
      }
      await context.sync();

      for (let i = 0; i < tables.items.length; i++) {
        const table = tables.items[i];
        const values = table.values;
        parts.push("");
        parts.push(`=== TABLE ${i + 1} (${values.length} rows x ${values[0]?.length || 0} cols) ===`);
        for (let r = 0; r < values.length; r++) {
          const label = r === 0 ? "HEADER" : `ROW ${r}`;
          const cells = values[r].map((cell, c) => `[R${r}C${c}]${cell || "(empty)"}`).join(" | ");
          parts.push(`  ${label}: ${cells}`);
        }
      }
    }

    return parts.join("\n");
  });
}

export async function applyAIResponse(response: string): Promise<boolean> {
  // Extract JSON block from AI response
  const jsonMatch = response.match(/```json\s*([\s\S]*?)\s*```/);
  if (!jsonMatch) {
    return false; // No document modifications, just conversation
  }

  return Word.run(async (context) => {
    const body = context.document.body;
    const tables = body.tables;
    tables.load("items");
    await context.sync();

    const instructions: WriteInstruction[] = JSON.parse(jsonMatch[1]);
    await applyInstructions(context, instructions, tables);
    await context.sync();
    return true;
  });
}

interface WriteInstruction {
  type: "table_cell" | "paragraph" | "replace" | "insert_after_paragraph";
  tableIndex?: number;
  row?: number;
  col?: number;
  paragraphIndex?: number;
  searchText?: string;
  value: string;
  style?: string;
}

async function applyInstructions(
  context: Word.RequestContext,
  instructions: WriteInstruction[],
  tables: Word.TableCollection
): Promise<void> {
  for (const inst of instructions) {
    if (inst.type === "table_cell" && inst.tableIndex !== undefined) {
      if (inst.tableIndex >= tables.items.length) continue;
      const table = tables.items[inst.tableIndex];
      if (inst.row === undefined || inst.col === undefined) continue;
      table.load("rowCount");
      await context.sync();
      if (inst.row >= table.rowCount) continue;
      const cell = table.getCell(inst.row, inst.col);
      cell.body.clear();
      cell.body.insertText(inst.value, Word.InsertLocation.start);
    } else if (inst.type === "replace" && inst.searchText) {
      const results = context.document.body.search(inst.searchText, {
        matchCase: false,
        matchWholeWord: false,
      });
      results.load("items");
      await context.sync();
      if (results.items.length > 0) {
        results.items[0].insertText(inst.value, Word.InsertLocation.replace);
      }
    } else if (inst.type === "insert_after_paragraph" && inst.paragraphIndex !== undefined) {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items");
      await context.sync();
      if (inst.paragraphIndex < paragraphs.items.length) {
        const newPara = paragraphs.items[inst.paragraphIndex].insertParagraph(
          inst.value,
          Word.InsertLocation.after
        );
        if (inst.style) {
          newPara.style = inst.style;
        }
      }
    } else if (inst.type === "paragraph") {
      const newPara = context.document.body.insertParagraph(
        inst.value,
        Word.InsertLocation.end
      );
      if (inst.style) {
        newPara.style = inst.style;
      }
    }
  }
}
