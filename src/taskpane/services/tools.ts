import { readDocumentStructure } from "./wordDocument";

/**
 * Tool Architecture:
 * - 7 dedicated tools for the most common operations (reliable, undo-aware)
 * - 1 power tool (execute_word_js) for EVERYTHING else
 *
 * This mirrors Claude Code's approach: a few structured tools + one general-purpose executor.
 */

export const TOOL_DEFINITIONS = [
  // === READING ===
  {
    type: "function" as const,
    function: {
      name: "read_document",
      description:
        "Read the current Word document structure including all paragraphs (with styles, alignment) and tables (with cell values, width). ALWAYS call this first before making changes.",
      parameters: { type: "object", properties: {}, required: [] },
    },
  },

  // === WRITING TEXT ===
  {
    type: "function" as const,
    function: {
      name: "batch_write",
      description:
        "Write multiple paragraphs to the document at once. Use this for any text content creation. Each paragraph can have its own style.",
      parameters: {
        type: "object",
        properties: {
          paragraphs: {
            type: "array",
            items: {
              type: "object",
              properties: {
                text: { type: "string", description: "Text content" },
                style: {
                  type: "string",
                  description:
                    "Paragraph style: 'Normal', 'Heading 1', 'Heading 2', 'Heading 3', 'Title', 'Subtitle', 'List Bullet', 'List Number'",
                },
              },
              required: ["text"],
            },
            description: "Array of paragraphs to append at end of document",
          },
          insertAfterIndex: {
            type: "number",
            description:
              "If provided, insert after this 0-based paragraph index instead of appending at end",
          },
        },
        required: ["paragraphs"],
      },
    },
  },

  // === REPLACE TEXT ===
  {
    type: "function" as const,
    function: {
      name: "replace_text",
      description: "Find and replace text in the document.",
      parameters: {
        type: "object",
        properties: {
          searchText: { type: "string", description: "Text to search for" },
          replaceWith: { type: "string", description: "Replacement text" },
          replaceAll: { type: "boolean", description: "Replace all occurrences (default: false)" },
        },
        required: ["searchText", "replaceWith"],
      },
    },
  },

  // === TABLE CREATION ===
  {
    type: "function" as const,
    function: {
      name: "insert_table",
      description: "Create a new table with headers and data rows. Appended at end of document.",
      parameters: {
        type: "object",
        properties: {
          headers: {
            type: "array",
            items: { type: "string" },
            description: "Column header labels",
          },
          rows: {
            type: "array",
            items: { type: "array", items: { type: "string" } },
            description: "2D array of row data",
          },
          style: {
            type: "string",
            description: "Word table style (default: 'Grid Table 1 Light')",
          },
        },
        required: ["headers", "rows"],
      },
    },
  },

  // === TABLE CELL EDITING ===
  {
    type: "function" as const,
    function: {
      name: "write_table_cells",
      description: "Write values to specific cells in an existing table.",
      parameters: {
        type: "object",
        properties: {
          changes: {
            type: "array",
            items: {
              type: "object",
              properties: {
                tableIndex: { type: "number", description: "0-based table index" },
                row: { type: "number", description: "0-based row index" },
                col: { type: "number", description: "0-based column index" },
                value: { type: "string", description: "Content to write" },
              },
              required: ["tableIndex", "row", "col", "value"],
            },
          },
        },
        required: ["changes"],
      },
    },
  },

  // === CLEAR DOCUMENT ===
  {
    type: "function" as const,
    function: {
      name: "clear_document",
      description: "Clear all content from the document. Use with caution!",
      parameters: { type: "object", properties: {}, required: [] },
    },
  },

  // === OOXML (for complex formatting) ===
  {
    type: "function" as const,
    function: {
      name: "insert_ooxml",
      description:
        "Insert Office Open XML (OOXML) at end of document. Use for complex formatted content that can't be done with plain text — e.g. colored text, mixed formatting in one paragraph, images via base64, complex tables with merged cells. Prefer batch_write for simple text.",
      parameters: {
        type: "object",
        properties: {
          ooxml: {
            type: "string",
            description: "OOXML string to insert",
          },
          description: {
            type: "string",
            description: "Brief description of what this inserts",
          },
        },
        required: ["ooxml", "description"],
      },
    },
  },

  // === DYNAMIC CODE EXECUTION (covers ALL Word operations) ===
  {
    type: "function" as const,
    function: {
      name: "execute_word_js",
      description: `Execute arbitrary Word JavaScript API code. Use this for ANY operation not covered by other tools — formatting, table resizing, borders, cell shading, margins, headers/footers, page breaks, deleting content, column widths, font changes, images, styles, sections, etc.

The code runs inside Word.run(async (context) => { YOUR CODE HERE; await context.sync(); }).
You have access to 'context' and 'Word' objects.

Common patterns:
- Table width: tables.items[0].width = 300;
- Autofit: tables.items[0].autoFitWindow();
- Delete paragraph: paragraphs.items[5].delete();
- Delete table: tables.items[0].delete();
- Add table row: tables.items[0].addRows("End", 1, [["a","b","c"]]);
- Bold text: range.font.bold = true; range.font.size = 14;
- Alignment: paragraphs.items[0].alignment = Word.Alignment.centered;
- Page break: body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
- Header: context.document.sections.getFirst().getHeader("Primary").insertParagraph("Header Text", "Start");
- Footer: context.document.sections.getFirst().getFooter("Primary").insertParagraph("Footer", "Start");
- Borders: table.getBorder("Top").type = "Single"; table.getBorder("Top").color = "#000";
- Cell font: table.getCell(0,0).body.font.bold = true;

IMPORTANT: Always load collections before accessing items:
  const tables = context.document.body.tables; tables.load("items"); await context.sync();`,
      parameters: {
        type: "object",
        properties: {
          code: {
            type: "string",
            description: "JavaScript code to execute inside Word.run(). Must use context and Word objects.",
          },
          description: {
            type: "string",
            description: "Brief description of what this code does",
          },
        },
        required: ["code", "description"],
      },
    },
  },
];

// ===== UNDO SYSTEM =====
export interface UndoEntry {
  description: string;
  changes: Array<{
    type: string;
    location: string;
    oldValue: string;
    newValue: string;
  }>;
}

const undoStack: UndoEntry[] = [];
const MAX_UNDO = 20;

export function getUndoStack(): UndoEntry[] {
  return undoStack;
}

export function clearUndoStack(): void {
  undoStack.length = 0;
}

function pushUndo(description: string, changes: UndoEntry["changes"] = []): UndoEntry {
  const entry: UndoEntry = { description, changes };
  undoStack.push(entry);
  if (undoStack.length > MAX_UNDO) undoStack.shift();
  return entry;
}

// ===== TOOL EXECUTOR =====
export async function executeTool(
  name: string,
  args: Record<string, any>
): Promise<{ result: string; undoEntry?: UndoEntry }> {
  try {
    switch (name) {
      case "read_document": {
        try {
          const structure = await readDocumentStructure();
          return { result: structure };
        } catch {
          return { result: "(Could not read document — running outside Word)" };
        }
      }
      case "batch_write":
        return executeBatchWrite(args.paragraphs, args.insertAfterIndex);
      case "replace_text":
        return executeReplaceText(args.searchText, args.replaceWith, args.replaceAll);
      case "insert_table":
        return executeInsertTable(args.headers, args.rows, args.style);
      case "write_table_cells":
        return executeWriteTableCells(args.changes);
      case "clear_document":
        return executeClearDocument();
      case "insert_ooxml":
        return executeInsertOoxml(args.ooxml, args.description);
      case "execute_word_js":
        return executeWordJs(args.code, args.description);
      default:
        return { result: `Unknown tool: ${name}` };
    }
  } catch (err: any) {
    return { result: `Tool error: ${err.message}` };
  }
}

// ===== IMPLEMENTATIONS =====

async function executeBatchWrite(
  paragraphs: Array<{ text: string; style?: string }>,
  insertAfterIndex?: number
): Promise<{ result: string; undoEntry?: UndoEntry }> {
  return Word.run(async (context) => {
    const body = context.document.body;

    if (insertAfterIndex !== undefined) {
      // Insert after specific paragraph
      const existingParas = body.paragraphs;
      existingParas.load("items");
      await context.sync();

      if (insertAfterIndex >= existingParas.items.length) {
        return { result: `Paragraph index ${insertAfterIndex} out of range (max: ${existingParas.items.length - 1})` };
      }

      let anchor = existingParas.items[insertAfterIndex];
      for (const p of paragraphs) {
        const newPara = anchor.insertParagraph(p.text, Word.InsertLocation.after);
        if (p.style) newPara.style = p.style;
        anchor = newPara;
      }
    } else {
      // Append at end
      for (const p of paragraphs) {
        const newPara = body.insertParagraph(p.text, Word.InsertLocation.end);
        if (p.style) newPara.style = p.style;
      }
    }
    await context.sync();

    const entry = pushUndo(`Wrote ${paragraphs.length} paragraph(s)`, [
      { type: "batch_write", location: insertAfterIndex !== undefined ? `After P${insertAfterIndex}` : "End", oldValue: "", newValue: `${paragraphs.length} paragraphs` },
    ]);
    return { result: `Wrote ${paragraphs.length} paragraph(s)`, undoEntry: entry };
  });
}

async function executeReplaceText(
  searchText: string,
  replaceWith: string,
  replaceAll?: boolean
): Promise<{ result: string; undoEntry?: UndoEntry }> {
  return Word.run(async (context) => {
    const results = context.document.body.search(searchText, {
      matchCase: false,
      matchWholeWord: false,
    });
    results.load("items");
    await context.sync();

    if (results.items.length === 0) {
      return { result: `Text "${searchText}" not found in document body. Note: search only covers body text, not tables or headers/footers. Try read_document to verify the exact text.` };
    }

    const count = replaceAll ? results.items.length : 1;
    for (let i = 0; i < count; i++) {
      results.items[i].insertText(replaceWith, Word.InsertLocation.replace);
    }
    await context.sync();

    const entry = pushUndo(`Replaced "${searchText}" (${count}x)`, [
      { type: "replace", location: "Document body", oldValue: searchText, newValue: replaceWith },
    ]);
    return { result: `Replaced ${count} occurrence(s) of "${searchText}"`, undoEntry: entry };
  });
}

async function executeInsertTable(
  headers: string[],
  rows: string[][],
  style?: string
): Promise<{ result: string; undoEntry?: UndoEntry }> {
  return Word.run(async (context) => {
    const body = context.document.body;
    const allRows = [headers, ...rows];
    const rowCount = allRows.length;
    const colCount = headers.length;

    const paddedRows = allRows.map((row) => {
      const padded = [...row];
      while (padded.length < colCount) padded.push("");
      return padded.slice(0, colCount);
    });

    const table = body.insertTable(rowCount, colCount, Word.InsertLocation.end, paddedRows);

    try {
      table.style = style || "Grid Table 1 Light";
    } catch {
      // Style might not exist
    }

    // Bold headers
    for (let c = 0; c < colCount; c++) {
      table.getCell(0, c).body.font.bold = true;
    }
    await context.sync();

    const entry = pushUndo(`Inserted ${rowCount}x${colCount} table`, [
      { type: "insert_table", location: "End of document", oldValue: "", newValue: `${rowCount}x${colCount}` },
    ]);
    return { result: `Created table: ${colCount} cols × ${rows.length} rows`, undoEntry: entry };
  });
}

async function executeWriteTableCells(
  changes: Array<{ tableIndex: number; row: number; col: number; value: string }>
): Promise<{ result: string; undoEntry?: UndoEntry }> {
  return Word.run(async (context) => {
    const tables = context.document.body.tables;
    tables.load("items");
    await context.sync();

    const undoChanges: UndoEntry["changes"] = [];
    let appliedCount = 0;
    const skipped: string[] = [];

    for (const change of changes) {
      if (change.tableIndex >= tables.items.length) {
        skipped.push(`Table ${change.tableIndex} does not exist (only ${tables.items.length} tables)`);
        continue;
      }
      const table = tables.items[change.tableIndex];
      table.load("rowCount,values");
      await context.sync();

      if (change.row >= table.rowCount) {
        skipped.push(`Table ${change.tableIndex + 1} row ${change.row} out of range (max: ${table.rowCount - 1})`);
        continue;
      }
      const oldValue = table.values[change.row]?.[change.col] || "";

      const cell = table.getCell(change.row, change.col);
      cell.body.clear();
      cell.body.insertText(change.value, Word.InsertLocation.start);
      appliedCount++;

      undoChanges.push({
        type: "table_cell",
        location: `Table ${change.tableIndex + 1}, Row ${change.row}, Col ${change.col}`,
        oldValue,
        newValue: change.value,
      });
    }
    await context.sync();

    const entry = pushUndo(`Modified ${appliedCount} table cell(s)`, undoChanges);
    const skippedMsg = skipped.length > 0 ? ` Skipped: ${skipped.join("; ")}` : "";
    return { result: `Wrote ${appliedCount} cell(s).${skippedMsg}`, undoEntry: entry };
  });
}

async function executeClearDocument(): Promise<{ result: string; undoEntry?: UndoEntry }> {
  return Word.run(async (context) => {
    context.document.body.clear();
    await context.sync();
    const entry = pushUndo("Cleared document");
    return { result: "Document cleared", undoEntry: entry };
  });
}

async function executeInsertOoxml(
  ooxml: string,
  description: string
): Promise<{ result: string; undoEntry?: UndoEntry }> {
  return Word.run(async (context) => {
    context.document.body.insertOoxml(ooxml, Word.InsertLocation.end);
    await context.sync();
    const entry = pushUndo(description, [
      { type: "insert_ooxml", location: "End of document", oldValue: "", newValue: description },
    ]);
    return { result: `Inserted: ${description}`, undoEntry: entry };
  });
}

// Forbidden patterns in execute_word_js to prevent code injection
const FORBIDDEN_PATTERNS = [
  /\beval\s*\(/,
  /\bimport\s*\(/,
  /\brequire\s*\(/,
  /\bfetch\s*\(/,
  /\bXMLHttpRequest\b/,
  /\blocalStorage\b/,
  /\bsessionStorage\b/,
  /\bdocument\.(cookie|write|location)/,
  /\bwindow\.(open|location)/,
  /\bnew\s+Worker\b/,
];

async function executeWordJs(
  code: string,
  description: string
): Promise<{ result: string; undoEntry?: UndoEntry }> {
  // Validate code before execution
  for (const pattern of FORBIDDEN_PATTERNS) {
    if (pattern.test(code)) {
      return { result: `Error: Code contains forbidden pattern (${pattern.source}). Only Word API operations are allowed.` };
    }
  }

  try {
    const result = await Word.run(async (context) => {
      const fn = new Function(
        "context",
        "Word",
        `return (async () => { ${code} })();`
      );
      const res = await fn(context, Word);
      await context.sync();
      return typeof res === "string" ? res : (res !== undefined ? JSON.stringify(res) : "Done");
    });

    const entry = pushUndo(description, [
      { type: "execute_word_js", location: "Document", oldValue: "", newValue: description },
    ]);
    return { result: result || `Done: ${description}`, undoEntry: entry };
  } catch (err: any) {
    return { result: `Execution error: ${err.message}` };
  }
}

// ===== UNDO =====
export async function undoLastOperation(): Promise<string> {
  const entry = undoStack.pop();
  if (!entry) return "Nothing to undo";

  // Only table_cell and replace have true undo capability
  const reversibleChanges = entry.changes.filter(
    (c) => c.type === "table_cell" || c.type === "replace"
  );

  if (reversibleChanges.length === 0) {
    return `Cannot fully undo: "${entry.description}" — use Ctrl+Z in Word for complete undo`;
  }

  return Word.run(async (context) => {
    for (const change of reversibleChanges) {
      if (change.type === "table_cell") {
        const match = change.location.match(/Table (\d+), Row (\d+), Col (\d+)/);
        if (!match) continue;
        const tableIndex = parseInt(match[1]) - 1;
        const row = parseInt(match[2]);
        const col = parseInt(match[3]);

        const tables = context.document.body.tables;
        tables.load("items");
        await context.sync();

        if (tableIndex < tables.items.length) {
          const cell = tables.items[tableIndex].getCell(row, col);
          cell.body.clear();
          if (change.oldValue) {
            cell.body.insertText(change.oldValue, Word.InsertLocation.start);
          }
        }
      } else if (change.type === "replace") {
        const results = context.document.body.search(change.newValue, {
          matchCase: false,
          matchWholeWord: false,
        });
        results.load("items");
        await context.sync();
        if (results.items.length > 0) {
          results.items[0].insertText(change.oldValue, Word.InsertLocation.replace);
        }
      }
    }
    await context.sync();
    return `Undone: ${entry.description}`;
  });
}
