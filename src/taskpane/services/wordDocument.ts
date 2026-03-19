export async function readDocumentStructure(): Promise<string> {
  return Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");

    const tables = body.tables;
    tables.load("items");

    await context.sync();

    const parts: string[] = [];
    parts.push("=== DOCUMENT TEXT ===");
    parts.push(body.text);

    if (tables.items.length > 0) {
      for (let i = 0; i < tables.items.length; i++) {
        const table = tables.items[i];
        table.load("rowCount");
        table.load("values");
      }
      await context.sync();

      for (let i = 0; i < tables.items.length; i++) {
        const table = tables.items[i];
        parts.push(`\n=== TABLE ${i + 1} ===`);
        const values = table.values;
        for (let r = 0; r < values.length; r++) {
          parts.push(values[r].join(" | "));
        }
      }
    }

    return parts.join("\n");
  });
}

export async function applyAIResponse(response: string): Promise<void> {
  return Word.run(async (context) => {
    const body = context.document.body;
    const tables = body.tables;
    tables.load("items");
    await context.sync();

    // Try to parse structured JSON response
    const jsonMatch = response.match(/```json\s*([\s\S]*?)\s*```/);
    if (jsonMatch) {
      const instructions = JSON.parse(jsonMatch[1]);
      await applyInstructions(context, instructions, tables);
    } else {
      // Fallback: insert response as text at end of document
      body.insertParagraph(response, Word.InsertLocation.end);
    }

    await context.sync();
  });
}

interface WriteInstruction {
  type: "table_cell" | "paragraph" | "replace";
  tableIndex?: number;
  row?: number;
  col?: number;
  paragraphIndex?: number;
  searchText?: string;
  value: string;
}

async function applyInstructions(
  context: Word.RequestContext,
  instructions: WriteInstruction[],
  tables: Word.TableCollection
): Promise<void> {
  for (const inst of instructions) {
    if (inst.type === "table_cell" && inst.tableIndex !== undefined) {
      const table = tables.items[inst.tableIndex];
      const cell = table.getCell(inst.row!, inst.col!);
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
    } else if (inst.type === "paragraph") {
      const body = context.document.body;
      body.insertParagraph(inst.value, Word.InsertLocation.end);
    }
  }
}
