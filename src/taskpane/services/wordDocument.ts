export async function readDocumentStructure(): Promise<string> {
  return Word.run(async (context) => {
    const body = context.document.body;

    // Load paragraphs with style info
    const paragraphs = body.paragraphs;
    paragraphs.load("items");
    await context.sync();

    for (const para of paragraphs.items) {
      para.load("text,style,alignment");
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
        table.load("rowCount,values,width");
      }
      await context.sync();

      for (let i = 0; i < tables.items.length; i++) {
        const table = tables.items[i];
        const values = table.values;
        parts.push("");
        parts.push(
          `=== TABLE ${i + 1} (${values.length} rows x ${values[0]?.length || 0} cols, width: ${table.width}pt) ===`
        );
        for (let r = 0; r < values.length; r++) {
          const label = r === 0 ? "HEADER" : `ROW ${r}`;
          const cells = values[r]
            .map((cell, c) => `[R${r}C${c}]${cell || "(empty)"}`)
            .join(" | ");
          parts.push(`  ${label}: ${cells}`);
        }
      }
    }

    return parts.join("\n");
  });
}
