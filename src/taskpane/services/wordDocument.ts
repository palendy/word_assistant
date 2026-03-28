/**
 * Reads the complete document structure including:
 * - Paragraphs with text, style, alignment, and font formatting (bold/italic/size/color)
 * - Tables with cell values, dimensions, and width
 * - Headers and footers
 * - Skips consecutive empty paragraphs to reduce token waste
 */
export async function readDocumentStructure(): Promise<string> {
  return Word.run(async (context) => {
    const body = context.document.body;

    // Load paragraphs
    const paragraphs = body.paragraphs;
    paragraphs.load("items");
    await context.sync();

    for (const para of paragraphs.items) {
      para.load("text,style,alignment");
      para.font.load("bold,italic,size,color,name");
    }

    // Load tables
    const tables = body.tables;
    tables.load("items");
    await context.sync();

    // Load sections for headers/footers
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();

    const parts: string[] = [];

    // === HEADERS & FOOTERS ===
    try {
      for (let s = 0; s < sections.items.length; s++) {
        const section = sections.items[s];
        const header = section.getHeader(Word.HeaderFooterType.primary);
        const footer = section.getFooter(Word.HeaderFooterType.primary);
        header.load("text");
        footer.load("text");
        await context.sync();

        const headerText = header.text?.trim();
        const footerText = footer.text?.trim();

        if (headerText || footerText) {
          parts.push(`=== SECTION ${s + 1} ===`);
          if (headerText) parts.push(`  HEADER: "${headerText}"`);
          if (footerText) parts.push(`  FOOTER: "${footerText}"`);
          parts.push("");
        }
      }
    } catch {
      // Headers/footers may not be accessible in all contexts
    }

    // === SUMMARY ===
    const nonEmptyCount = paragraphs.items.filter((p) => p.text.trim()).length;
    parts.push("=== DOCUMENT STRUCTURE ===");
    parts.push(`Paragraphs: ${paragraphs.items.length} total (${nonEmptyCount} with content)`);
    parts.push(`Tables: ${tables.items.length}`);
    parts.push("");

    // === PARAGRAPHS ===
    parts.push("=== PARAGRAPHS ===");
    let consecutiveEmpty = 0;

    for (let i = 0; i < paragraphs.items.length; i++) {
      const p = paragraphs.items[i];
      const text = p.text.trim();

      if (!text) {
        consecutiveEmpty++;
        // Show first empty, then collapse subsequent ones
        if (consecutiveEmpty === 1) {
          parts.push(`[P${i}] (empty)`);
        } else if (consecutiveEmpty === 2) {
          parts.push(`  ... (more empty paragraphs)`);
        }
        continue;
      }

      consecutiveEmpty = 0;

      // Build formatting tags
      const fmt: string[] = [];
      try {
        if (p.font.bold) fmt.push("bold");
        if (p.font.italic) fmt.push("italic");
        if (p.font.size) fmt.push(`${p.font.size}pt`);
        if (p.font.color && p.font.color !== "#000000" && p.font.color !== "Automatic") {
          fmt.push(`color:${p.font.color}`);
        }
        if (p.font.name) fmt.push(p.font.name);
      } catch {
        // Font properties may not be loaded in some edge cases
      }

      const fmtStr = fmt.length > 0 ? ` [${fmt.join(", ")}]` : "";
      parts.push(`[P${i}] "${text}" | style: ${p.style} | align: ${p.alignment}${fmtStr}`);
    }

    // === TABLES ===
    if (tables.items.length > 0) {
      for (let i = 0; i < tables.items.length; i++) {
        const table = tables.items[i];
        table.load("rowCount,values,width,style");
      }
      await context.sync();

      for (let i = 0; i < tables.items.length; i++) {
        const table = tables.items[i];
        const values = table.values;
        parts.push("");
        parts.push(
          `=== TABLE ${i + 1} (${values.length} rows x ${values[0]?.length || 0} cols, width: ${Math.round(table.width)}pt, style: ${table.style || "none"}) ===`
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
