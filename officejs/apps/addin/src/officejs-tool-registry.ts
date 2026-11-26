import { OfficeJsWordAdapter } from "./officejs-word-adapter.ts";
import type { OfficeToolDescriptor } from "./contracts";
import i18n from "./i18n.ts";

type ToolArguments = Record<string, unknown>;

const expectedSelectionHashSchema = {
  type: "string",
  pattern: "^[0-9a-f]{64}$",
  description: "Opaque selection hash returned by get_selection or select_exact_text. The write is rejected if the selection changed.",
};

const descriptors: OfficeToolDescriptor[] = [
  {
    name: "get_selection",
    description: "Read the current Word selection and return the selectionHash required by selection-based write tools.",
    isWriteOperation: false,
    parameterSchema: { type: "object", properties: {}, required: [] },
  },
  {
    name: "select_exact_text",
    description: "Select text only when one unique, case-sensitive exact match exists and return its selectionHash; refuses zero or ambiguous matches. Changes the visible Word selection but not document content.",
    // Selection is visible, shared host state. Treat it as a write for MCP
    // approval and multi-document targeting even though document text is unchanged.
    isWriteOperation: true,
    parameterSchema: {
      type: "object",
      properties: {
        text: { type: "string", minLength: 1, maxLength: 255 },
      },
      required: ["text"],
      additionalProperties: false,
    },
  },
  {
    name: "search_text",
    description: "Search for text in the current Word document.",
    isWriteOperation: false,
    parameterSchema: {
      type: "object",
      properties: { keyword: { type: "string" } },
      required: ["keyword"],
    },
  },
  {
    name: "get_doc_overview",
    description: "Read the document paragraph count and opening preview. Use first for a quick orientation before reading specific ranges.",
    isWriteOperation: false,
    parameterSchema: { type: "object", properties: {}, required: [] },
  },
  {
    name: "insert_text_at_end",
    description: "Insert text at the end of the current Word document.",
    isWriteOperation: true,
    parameterSchema: {
      type: "object",
      properties: { text: { type: "string" } },
      required: ["text"],
    },
  },
  {
    name: "read_paragraphs",
    description: "Read a range of Word paragraphs by 1-based inclusive start and end indices.",
    isWriteOperation: false,
    parameterSchema: {
      type: "object",
      properties: {
        start: { type: "integer", minimum: 1, description: "1-based starting paragraph index." },
        end: { type: "integer", minimum: 1, description: "1-based ending paragraph index." },
      },
      required: ["start", "end"],
    },
  },
  {
    name: "read_large_chunk",
    description: "Read up to 50 paragraphs starting at a 1-based start_paragraph index.",
    isWriteOperation: false,
    parameterSchema: {
      type: "object",
      properties: {
        start_paragraph: { type: "integer", minimum: 1, description: "1-based starting paragraph index." },
      },
      required: ["start_paragraph"],
    },
  },
  {
    name: "build_semantic_map",
    description: "Build a full-document structural map from paragraph chunks for long documents.",
    isWriteOperation: false,
    parameterSchema: { type: "object", properties: {}, required: [] },
  },
  {
    name: "read_comments",
    description: "Read comments in the current document.",
    isWriteOperation: false,
    parameterSchema: { type: "object", properties: {}, required: [] },
  },
  {
    name: "read_bookmarks",
    description: "Read bookmarks and their range text in the current document.",
    isWriteOperation: false,
    parameterSchema: { type: "object", properties: {}, required: [] },
  },
  {
    name: "apply_style",
    description: "Apply a style to a 1-based paragraph_index. Use the exact localized style_name from list_styles (e.g. 'Heading 1' or 'Nagłówek 1').",
    isWriteOperation: true,
    parameterSchema: {
      type: "object",
      properties: {
        paragraph_index: { type: "integer", minimum: 1, description: "1-based paragraph index." },
        style_name: { type: "string", description: "Exact localized style name from list_styles." },
      },
      required: ["paragraph_index", "style_name"],
    },
  },
  {
    name: "replace_paragraph",
    description: "Replace the text of a paragraph by its 1-based paragraph_index.",
    isWriteOperation: true,
    parameterSchema: {
      type: "object",
      properties: {
        paragraph_index: { type: "integer", minimum: 1, description: "1-based paragraph index." },
        new_text: { type: "string" },
      },
      required: ["paragraph_index", "new_text"],
    },
  },
  {
    name: "delete_paragraph",
    description: "Delete a paragraph by its 1-based paragraph_index.",
    isWriteOperation: true,
    parameterSchema: {
      type: "object",
      properties: {
        paragraph_index: {
          type: "integer",
          minimum: 1,
          description: "1-based index of the paragraph to delete.",
        },
      },
      required: ["paragraph_index"],
    },
  },
  {
    name: "insert_at_cursor",
    description: "Insert text at the current Word cursor or selection.",
    isWriteOperation: true,
    parameterSchema: {
      type: "object",
      properties: { text: { type: "string" }, expected_selection_hash: expectedSelectionHashSchema },
      required: ["text", "expected_selection_hash"],
    },
  },
  {
    name: "insert_hyperlink",
    description: "Insert a clickable hyperlink at the cursor or over the current selection.",
    isWriteOperation: true,
    parameterSchema: {
      type: "object",
      properties: {
        text: { type: "string", description: "Visible link text to insert at the cursor or replace selection." },
        url: { type: "string", description: "Target URL (http:, https:, or mailto:)." },
        expected_selection_hash: expectedSelectionHashSchema,
      },
      required: ["text", "url", "expected_selection_hash"],
    },
  },
  {
    name: "add_comment",
    description: "Add a comment to the current selection.",
    isWriteOperation: true,
    parameterSchema: {
      type: "object",
      properties: { text: { type: "string" }, expected_selection_hash: expectedSelectionHashSchema },
      required: ["text", "expected_selection_hash"],
    },
  },
  {
    name: "find_replace",
    description: "Find and replace text in the document.",
    isWriteOperation: true,
    parameterSchema: {
      type: "object",
      properties: { find: { type: "string" }, replace: { type: "string" } },
      required: ["find", "replace"],
    },
  },
  {
    name: "format_paragraph",
    description: "Set alignment and spacing (in points) for a paragraph by its 1-based paragraph_index.",
    isWriteOperation: true,
    parameterSchema: {
      type: "object",
      properties: {
        paragraph_index: { type: "integer", minimum: 1, description: "1-based paragraph index." },
        alignment: { type: "string", enum: ["Left", "Centered", "Right", "Justified"] },
        space_before: { type: "number", description: "Spacing before paragraph in points (pt), e.g. 12." },
        space_after: { type: "number", description: "Spacing after paragraph in points (pt), e.g. 6." },
      },
      required: ["paragraph_index"],
    },
  },
  {
    name: "format_text",
    description: "Set font formatting for the current selection.",
    isWriteOperation: true,
    parameterSchema: {
      type: "object",
      properties: {
        bold: { type: "boolean" },
        italic: { type: "boolean" },
        underline: { type: "boolean" },
        font_name: { type: "string" },
        font_size: { type: "number", description: "Font size in points (pt), e.g. 11." },
        color: { type: "string", description: "Hex color (#RRGGBB), e.g. '#000000'." },
        expected_selection_hash: expectedSelectionHashSchema,
      },
      required: ["expected_selection_hash"],
    },
  },
  {
    name: "insert_page_break",
    description: "Insert a page break at the current cursor.",
    isWriteOperation: true,
    parameterSchema: {
      type: "object",
      properties: { expected_selection_hash: expectedSelectionHashSchema },
      required: ["expected_selection_hash"],
    },
  },
  {
    name: "read_table",
    description: "Read a specified Word table as two-dimensional cell data. table_index is 1-based (first table is 1).",
    isWriteOperation: false,
    parameterSchema: {
      type: "object",
      properties: {
        table_index: {
          type: "integer",
          minimum: 1,
          description: "1-based index of the table in the document (first table is 1, not 0).",
        },
      },
      required: ["table_index"],
    },
  },
  {
    name: "insert_table",
    description: "Insert a Word table at the end of the document.",
    isWriteOperation: true,
    parameterSchema: {
      type: "object",
      properties: {
        rows: { type: "integer" },
        columns: { type: "integer" },
        header_row: { type: "array", items: { type: "string" } },
      },
      required: ["rows", "columns"],
    },
  },
  {
    name: "table_set_cell",
    description: "Set text in a Word table cell (1-based table_index, row, and column).",
    isWriteOperation: true,
    parameterSchema: {
      type: "object",
      properties: {
        table_index: { type: "integer", minimum: 1 },
        row: { type: "integer", minimum: 1 },
        column: { type: "integer", minimum: 1 },
        text: { type: "string" },
      },
      required: ["table_index", "row", "column", "text"],
    },
  },
  {
    name: "edit_table_structure",
    description: "Insert or delete table rows/columns, or merge and split cells (1-based table_index, row, and column).",
    isWriteOperation: true,
    parameterSchema: {
      type: "object",
      properties: {
        table_index: { type: "integer", minimum: 1 },
        action: { type: "string", enum: ["insert_row", "delete_row", "insert_column", "delete_column", "merge_cells", "split_cell"] },
        row: { type: "integer", minimum: 1 },
        column: { type: "integer", minimum: 1 },
        end_row: { type: "integer", minimum: 1 },
        end_column: { type: "integer", minimum: 1 },
        count: { type: "integer", minimum: 1 },
      },
      required: ["table_index", "action"],
    },
  },
  {
    name: "format_table",
    description: "Set table style, header rows, alignment, shading, borders, and AutoFit.",
    isWriteOperation: true,
    parameterSchema: {
      type: "object",
      properties: {
        table_index: { type: "integer" },
        style_name: { type: "string" },
        header_rows: { type: "integer" },
        alignment: { type: "string", enum: ["Left", "Centered", "Right"] },
        cell_alignment: { type: "string", enum: ["Left", "Centered", "Right", "Justified"] },
        vertical_alignment: { type: "string", enum: ["Top", "Center", "Bottom"] },
        shading_color: { type: "string" },
        border_color: { type: "string", description: "Hex color (#RRGGBB), e.g. '#000000'." },
        border_width: { type: "number", description: "Border width in points (pt), e.g. 0.5 or 1." },
        autofit: { type: "string", enum: ["window", "content", "fixed"] },
      },
      required: ["table_index"],
    },
  },
  {
    name: "get_document_outline",
    description: "Read the document outline from heading paragraphs.",
    isWriteOperation: false,
    parameterSchema: { type: "object", properties: {}, required: [] },
  },
  {
    name: "clear_formatting",
    description: "Clear direct character formatting, paragraph spacing/indents, or reset paragraphs to Normal style across the document or current selection.",
    isWriteOperation: true,
    parameterSchema: {
      type: "object",
      properties: {
        scope: {
          type: "string",
          enum: ["document", "selection"],
          description: "Target scope: 'document' (default) or 'selection' (requires expected_selection_hash).",
        },
        clear_font: { type: "boolean", description: "Clear bold, italic, underline, strikethrough, and highlight (default true)." },
        clear_paragraph_spacing: { type: "boolean", description: "Reset paragraph spacing and indents to 0 (default true)." },
        reset_to_normal: { type: "boolean", description: "Reset paragraph style to built-in Normal style (default false)." },
        expected_selection_hash: expectedSelectionHashSchema,
      },
      required: [],
    },
  },
  {
    name: "insert_image",
    description: "Insert a Base64 image at the end of the document.",
    isWriteOperation: true,
    parameterSchema: {
      type: "object",
      properties: { base64: { type: "string" }, alt_text: { type: "string" } },
      required: ["base64"],
    },
  },
  {
    name: "format_list",
    description: "Format the currently selected paragraphs as a bulleted or numbered list (select target text with select_exact_text first).",
    isWriteOperation: true,
    parameterSchema: {
      type: "object",
      properties: {
        list_type: { type: "string", enum: ["bullet", "number"] },
        expected_selection_hash: expectedSelectionHashSchema,
      },
      required: ["list_type", "expected_selection_hash"],
    },
  },
  {
    name: "get_active_page",
    description: "Detect the 1-based page number of the current cursor/selection, visible viewport pages, and total page count.",
    isWriteOperation: false,
    parameterSchema: { type: "object", properties: {}, required: [] },
  },
  {
    name: "get_page_preview",
    description: "Render a visual PNG preview of a Word document page. Omit page to render the current active page immediately without a separate lookup.",
    isWriteOperation: false,
    parameterSchema: {
      type: "object",
      properties: {
        page: { type: "integer", minimum: 1, description: "1-based page number to render. Defaults to the current active page." },
        dpi: { type: "integer", minimum: 72, maximum: 200, description: "Render resolution in DPI (72-200, default 120)." },
      },
      required: [],
    },
  },
  {
    name: "get_page_setup",
    description: "Read current page setup (margins, orientation, paper size, dimensions, header/footer distances, and section flags) for all document sections, in points.",
    isWriteOperation: false,
    parameterSchema: { type: "object", properties: {}, required: [] },
  },
  {
    name: "page_setup",
    description: "Set margins and orientation for document sections, in points.",
    isWriteOperation: true,
    parameterSchema: {
      type: "object",
      properties: {
        margin_top: { type: "number" },
        margin_bottom: { type: "number" },
        margin_left: { type: "number" },
        margin_right: { type: "number" },
        orientation: { type: "string", enum: ["Portrait", "Landscape"] },
        paper_size: { type: "string", enum: ["A3", "A4", "A5", "Letter", "Legal", "Custom"] },
        page_width: { type: "number" },
        page_height: { type: "number" },
        gutter: { type: "number" },
        header_distance: { type: "number" },
        footer_distance: { type: "number" },
        different_first_page: { type: "boolean" },
        odd_even_pages: { type: "boolean" },
        mirror_margins: { type: "boolean" },
      },
      required: [],
    },
  },
  {
    name: "read_header_footer",
    description: "Read header and footer text for document sections (1-based section_index, optional kind: primary, first, even).",
    isWriteOperation: false,
    parameterSchema: {
      type: "object",
      properties: {
        section_index: { type: "integer", minimum: 1, description: "Optional 1-based section index. Omit to read all sections." },
        kind: { type: "string", enum: ["primary", "first", "even"], description: "Optional header/footer kind. Omit to read all kinds." },
      },
      required: [],
    },
  },
  {
    name: "header_footer",
    description: "Set header text, footer text, or page numbers for a document section (1-based section_index, default 1) and kind (primary, first, or even).",
    isWriteOperation: true,
    parameterSchema: {
      type: "object",
      properties: {
        element: { type: "string", enum: ["header", "footer", "page_number"] },
        text: { type: "string" },
        section_index: { type: "integer" },
        kind: { type: "string", enum: ["primary", "first", "even"] },
        alignment: { type: "string", enum: ["Left", "Centered", "Right"] },
      },
      required: ["element"],
    },
  },
  {
    name: "update_toc",
    description: "Update an existing table of contents or insert one at the start of the document.",
    isWriteOperation: true,
    parameterSchema: {
      type: "object",
      properties: {
        action: { type: "string", enum: ["update", "insert"] },
        upper_heading_level: { type: "integer" },
        lower_heading_level: { type: "integer" },
        include_page_numbers: { type: "boolean" },
        right_align_page_numbers: { type: "boolean" },
        use_hyperlinks: { type: "boolean" },
      },
      required: ["action"],
    },
  },
  {
    name: "apply_precise_revision",
    description: "Apply only the changed hunks to the current selection as tracked Word revisions; fails safely if the selection changed.",
    isWriteOperation: true,
    parameterSchema: {
      type: "object",
      properties: {
        original: { type: "string" },
        revised: { type: "string" },
        expected_selection_hash: expectedSelectionHashSchema,
      },
      required: ["original", "revised", "expected_selection_hash"],
    },
  },
  {
    name: "ask_human",
    description: "Ask the Word user a question and wait for confirmation or input. Always provide 1 to 3 concise options when possible so the user can answer with one click.",
    isWriteOperation: false,
    parameterSchema: {
      type: "object",
      properties: {
        question: { type: "string" },
        options: {
          type: "array",
          items: { type: "string" },
          description: "Always provide 1 to 3 short suggested reply options so the user can answer with one click.",
        },
      },
      required: ["question"],
    },
  },
  {
    name: "list_styles",
    description: "List document styles (name, type, inUse, font color, size, font name).",
    isWriteOperation: false,
    parameterSchema: {
      type: "object",
      properties: {
        in_use_only: { type: "boolean" },
      },
      required: [],
    },
  },
  {
    name: "edit_style",
    description: "Modify font formatting (font_color, font_name, font_size, bold, italic) on a style or '*' for all styles in use.",
    isWriteOperation: true,
    parameterSchema: {
      type: "object",
      properties: {
        style_name: {
          type: "string",
          description: "Exact localized style name (e.g. 'Heading 1' or 'Nagłówek 1'), or '*' for all styles in use.",
        },
        font_color: { type: "string" },
        font_name: { type: "string" },
        font_size: { type: "number" },
        bold: { type: "boolean" },
        italic: { type: "boolean" },
        line_spacing: {
          type: "number",
          description: "Line spacing as a multiplier (e.g. 1.15, 1.5, 2.0) or exact points (>= 8).",
        },
        space_before: { type: "number", description: "Spacing before paragraph in points (pt), e.g. 12." },
        space_after: { type: "number", description: "Spacing after paragraph in points (pt), e.g. 6." },
      },
      required: ["style_name"],
    },
  },
  {
    name: "get_track_changes",
    description: "Read whether Word Track Changes is currently enabled and its mode.",
    isWriteOperation: false,
    parameterSchema: { type: "object", properties: {} },
  },
  {
    name: "set_track_changes",
    description: "Enable or disable Word Track Changes for the active document.",
    isWriteOperation: true,
    parameterSchema: {
      type: "object",
      properties: {
        enabled: { type: "boolean", description: "True to enable Track Changes, false to turn it off." },
        mine_only: { type: "boolean", description: "Optional: track only my own changes when enabled." },
      },
      required: ["enabled"],
    },
  },
  {
    name: "reply_to_comment",
    description: "Reply to an existing comment thread by its 1-based comment_index from read_comments.",
    isWriteOperation: true,
    parameterSchema: {
      type: "object",
      properties: {
        comment_index: { type: "integer", minimum: 1, description: "1-based index from read_comments." },
        text: { type: "string", description: "Reply text to add to the comment thread." },
      },
      required: ["comment_index", "text"],
    },
  },
  {
    name: "resolve_comment",
    description: "Mark an existing comment as resolved (or reopen it) by comment_index.",
    isWriteOperation: true,
    parameterSchema: {
      type: "object",
      properties: {
        comment_index: { type: "integer", minimum: 1, description: "1-based index from read_comments." },
        resolved: { type: "boolean", description: "True to resolve (default), false to reopen." },
      },
      required: ["comment_index"],
    },
  },
  {
    name: "delete_comment",
    description: "Delete an existing comment by its 1-based comment_index from read_comments.",
    isWriteOperation: true,
    parameterSchema: {
      type: "object",
      properties: {
        comment_index: { type: "integer", minimum: 1, description: "1-based index from read_comments." },
      },
      required: ["comment_index"],
    },
  },
  {
    name: "read_footnotes",
    description: "Read all footnotes in the active Word document.",
    isWriteOperation: false,
    parameterSchema: { type: "object", properties: {} },
  },
  {
    name: "insert_footnote",
    description: "Insert a footnote at the current cursor position or selection.",
    isWriteOperation: true,
    parameterSchema: {
      type: "object",
      properties: {
        text: { type: "string", description: "Footnote text to insert at the current cursor or selection." },
        expected_selection_hash: expectedSelectionHashSchema,
      },
      required: ["text", "expected_selection_hash"],
    },
  },
];

export class OfficeJsToolRegistry {
  private readonly word: OfficeJsWordAdapter;

  constructor(word: OfficeJsWordAdapter) {
    this.word = word;
  }

  list(): OfficeToolDescriptor[] {
    return descriptors.filter((descriptor) => this.word.supportsTool(descriptor.name)).map((descriptor) => ({
      ...descriptor,
      description: descriptor.description,
      // Tool schemas are JSON values. A JSON clone keeps WPS WebViews that do
      // not yet expose structuredClone from failing during task-pane startup.
      parameterSchema: JSON.parse(JSON.stringify(descriptor.parameterSchema)) as Record<string, unknown>,
    }));
  }

  async execute(name: string, args: ToolArguments = {}): Promise<unknown> {
    if (!this.word.supportsTool(name)) {
      throw new Error(i18n.t("taskpane.wordAdapter.errors.toolUnsupported", { name }));
    }
    switch (name) {
      case "get_selection":
        return this.word.getSelection();
      case "select_exact_text":
        return this.word.selectExactText(this.requireString(args, "text"));
      case "search_text":
        return this.word.searchText(this.requireString(args, "keyword"));
      case "get_doc_overview":
        return this.word.getDocumentOverview();
      case "insert_text_at_end":
        await this.word.insertTextAtEnd(this.requireString(args, "text"));
        return { inserted: true };
      case "read_paragraphs":
        return this.word.readParagraphs(this.requireInt(args, "start"), this.requireInt(args, "end"));
      case "read_large_chunk":
        return this.word.readLargeChunk(this.requireInt(args, "start_paragraph"));
      case "build_semantic_map":
        return this.word.buildSemanticMap();
      case "read_comments":
        return this.word.readComments();
      case "read_bookmarks":
        return this.word.readBookmarks();
      case "apply_style":
        await this.word.applyStyle(this.requireInt(args, "paragraph_index"), this.requireString(args, "style_name"));
        return { applied: true };
      case "replace_paragraph":
        await this.word.replaceParagraph(this.requireInt(args, "paragraph_index"), this.requireString(args, "new_text", true));
        return { replaced: true };
      case "delete_paragraph":
        await this.word.deleteParagraph(this.requireInt(args, "paragraph_index"));
        return { deleted: true };
      case "insert_at_cursor":
        await this.word.insertAtCursor(this.requireString(args, "text"), this.requireSelectionHash(args));
        return { inserted: true };
      case "insert_hyperlink":
        await this.word.insertHyperlink(
          this.requireString(args, "text"),
          this.requireString(args, "url"),
          this.requireSelectionHash(args),
        );
        return { inserted: true };
      case "add_comment":
        await this.word.addComment(this.requireString(args, "text"), this.requireSelectionHash(args));
        return { added: true };
      case "find_replace":
        return this.word.findReplace(this.requireString(args, "find"), this.requireString(args, "replace"));
      case "format_paragraph":
        await this.word.formatParagraph(
          this.requireInt(args, "paragraph_index"),
          this.optionalString(args, "alignment"),
          this.optionalNumber(args, "space_before"),
          this.optionalNumber(args, "space_after"),
        );
        return { formatted: true };
      case "format_text":
        await this.word.formatText({
          bold: this.optionalBoolean(args, "bold"),
          italic: this.optionalBoolean(args, "italic"),
          underline: this.optionalBoolean(args, "underline"),
          fontName: this.optionalString(args, "font_name"),
          fontSize: this.optionalNumber(args, "font_size"),
          color: this.optionalString(args, "color"),
        }, this.requireSelectionHash(args));
        return { formatted: true };
      case "insert_page_break":
        await this.word.insertPageBreak(this.requireSelectionHash(args));
        return { inserted: true };
      case "read_table":
        return this.word.readTable(this.requireInt(args, "table_index"));
      case "insert_table":
        await this.word.insertTable(
          this.requireInt(args, "rows"),
          this.requireInt(args, "columns"),
          this.optionalStringArray(args, "header_row"),
        );
        return { inserted: true };
      case "table_insert_row":
        await this.word.insertTableRow(
          this.requireInt(args, "table_index"),
          this.optionalInt(args, "after_row"),
        );
        return { inserted: true };
      case "table_set_cell":
        await this.word.setTableCell(
          this.requireInt(args, "table_index"),
          this.requireInt(args, "row"),
          this.requireInt(args, "column"),
          this.requireString(args, "text"),
        );
        return { updated: true };
      case "get_document_outline":
        return this.word.getDocumentOutline();
      case "read_clause":
        return this.word.readClause(this.requireString(args, "keyword"));
      case "extract_definitions":
        return this.word.extractDefinitions();
      case "check_cross_references":
        return this.word.checkCrossReferences();
      case "insert_clause_after":
        await this.word.insertClauseAfter(this.requireString(args, "anchor"), this.requireString(args, "text"));
        return { inserted: true };
      case "highlight_risk":
        return this.word.highlightRisk(this.requireString(args, "keyword"), this.optionalString(args, "color"));
      case "clear_formatting": {
        const scope = this.optionalString(args, "scope") ?? "document";
        return this.word.clearFormatting({
          scope,
          expectedSelectionHash: scope === "selection" ? this.requireSelectionHash(args) : undefined,
          clearFont: this.optionalBoolean(args, "clear_font"),
          clearParagraphSpacing: this.optionalBoolean(args, "clear_paragraph_spacing"),
          resetToNormal: this.optionalBoolean(args, "reset_to_normal"),
        });
      }
      case "validate_citation":
        return this.word.validateCitation(this.requireString(args, "citation"));
      case "insert_image":
        await this.word.insertImage(this.requireString(args, "base64"), this.optionalString(args, "alt_text"));
        return { inserted: true };
      case "format_list":
        await this.word.formatList(this.requireString(args, "list_type"), this.requireSelectionHash(args));
        return { formatted: true };
      case "get_active_page":
        return this.word.getActivePage();
      case "get_page_preview":
        return this.word.getPagePreview({
          page: this.optionalNumber(args, "page"),
          dpi: this.optionalNumber(args, "dpi"),
        });
      case "get_page_setup":
        return this.word.getPageSetup();
      case "page_setup":
        await this.word.pageSetup({
          marginTop: this.optionalNumber(args, "margin_top"),
          marginBottom: this.optionalNumber(args, "margin_bottom"),
          marginLeft: this.optionalNumber(args, "margin_left"),
          marginRight: this.optionalNumber(args, "margin_right"),
          orientation: this.optionalString(args, "orientation"),
          paperSize: this.optionalString(args, "paper_size"),
          pageWidth: this.optionalNumber(args, "page_width"),
          pageHeight: this.optionalNumber(args, "page_height"),
          gutter: this.optionalNumber(args, "gutter"),
          headerDistance: this.optionalNumber(args, "header_distance"),
          footerDistance: this.optionalNumber(args, "footer_distance"),
          differentFirstPage: this.optionalBoolean(args, "different_first_page"),
          oddEvenPages: this.optionalBoolean(args, "odd_even_pages"),
          mirrorMargins: this.optionalBoolean(args, "mirror_margins"),
        });
        return { updated: true };
      case "read_header_footer":
        return this.word.readHeaderFooter({
          sectionIndex: this.optionalNumber(args, "section_index"),
          kind: this.optionalString(args, "kind"),
        });
      case "header_footer":
        await this.word.headerFooter(
          this.requireString(args, "element"),
          this.optionalString(args, "text") ?? "",
          {
            sectionIndex: this.optionalInt(args, "section_index"),
            kind: this.optionalString(args, "kind"),
            alignment: this.optionalString(args, "alignment"),
          },
        );
        return { updated: true };
      case "edit_table_structure":
        await this.word.editTableStructure(
          this.requireInt(args, "table_index"),
          this.requireString(args, "action"),
          {
            row: this.optionalInt(args, "row"),
            column: this.optionalInt(args, "column"),
            endRow: this.optionalInt(args, "end_row"),
            endColumn: this.optionalInt(args, "end_column"),
            count: this.optionalInt(args, "count"),
          },
        );
        return { updated: true };
      case "format_table":
        await this.word.formatTable(this.requireInt(args, "table_index"), {
          styleName: this.optionalString(args, "style_name"),
          headerRows: this.optionalInt(args, "header_rows"),
          alignment: this.optionalString(args, "alignment"),
          cellAlignment: this.optionalString(args, "cell_alignment"),
          verticalAlignment: this.optionalString(args, "vertical_alignment"),
          shadingColor: this.optionalString(args, "shading_color"),
          borderColor: this.optionalString(args, "border_color"),
          borderWidth: this.optionalNumber(args, "border_width"),
          autofit: this.optionalString(args, "autofit"),
        });
        return { updated: true };
      case "update_toc":
        return this.word.updateToc(this.requireString(args, "action"), {
          upperHeadingLevel: this.optionalInt(args, "upper_heading_level"),
          lowerHeadingLevel: this.optionalInt(args, "lower_heading_level"),
          includePageNumbers: this.optionalBoolean(args, "include_page_numbers"),
          rightAlignPageNumbers: this.optionalBoolean(args, "right_align_page_numbers"),
          useHyperlinks: this.optionalBoolean(args, "use_hyperlinks"),
        });
      case "replace_exact_text":
        return this.word.findReplace(this.requireString(args, "find"), this.requireString(args, "replace"));
      case "apply_precise_revision":
        return {
          precise: await this.word.applyPreciseRevision(
            this.requireString(args, "original"),
            this.requireString(args, "revised", true),
            this.requireSelectionHash(args),
          ),
        };
      case "ask_human":
        return this.word.askHuman(
          this.requireString(args, "question"),
          this.optionalStringArray(args, "options"),
        );
      case "list_styles":
        return this.word.listStylesDetailed(this.optionalBoolean(args, "in_use_only") ?? true);
      case "edit_style":
        return this.word.editStyle(this.requireString(args, "style_name"), {
          fontColor: this.optionalString(args, "font_color"),
          fontName: this.optionalString(args, "font_name"),
          fontSize: this.optionalNumber(args, "font_size"),
          bold: this.optionalBoolean(args, "bold"),
          italic: this.optionalBoolean(args, "italic"),
          lineSpacing: this.optionalNumber(args, "line_spacing"),
          spaceBefore: this.optionalNumber(args, "space_before"),
          spaceAfter: this.optionalNumber(args, "space_after"),
        });
      case "get_track_changes":
        return this.word.getTrackChanges();
      case "set_track_changes": {
        const enabled = this.optionalBoolean(args, "enabled");
        if (enabled === undefined) throw new Error("enabled must be a boolean");
        return this.word.setTrackChanges(enabled, this.optionalBoolean(args, "mine_only") ?? false);
      }
      case "reply_to_comment":
        await this.word.replyToComment(this.requireInt(args, "comment_index"), this.requireString(args, "text"));
        return { replied: true };
      case "resolve_comment":
        await this.word.resolveComment(
          this.requireInt(args, "comment_index"),
          this.optionalBoolean(args, "resolved") ?? true,
        );
        return { resolved: this.optionalBoolean(args, "resolved") ?? true };
      case "delete_comment":
        await this.word.deleteComment(this.requireInt(args, "comment_index"));
        return { deleted: true };
      case "read_footnotes":
        return this.word.readFootnotes();
      case "insert_footnote":
        await this.word.insertFootnote(this.requireString(args, "text"), this.requireSelectionHash(args));
        return { inserted: true };
      default:
        throw new Error(`Office.js tool is not registered: ${name}`);
    }
  }

  private requireString(args: ToolArguments, name: string, allowEmpty = false): string {
    const value = args[name];
    if (typeof value !== "string" || (!allowEmpty && !value.trim())) {
      throw new Error(`${name} is required`);
    }
    return value;
  }

  private requireSelectionHash(args: ToolArguments): string {
    const value = this.requireString(args, "expected_selection_hash");
    if (!/^[0-9a-f]{64}$/u.test(value)) {
      throw new Error("expected_selection_hash must be a 64-character lowercase hexadecimal hash returned by get_selection or select_exact_text");
    }
    return value;
  }

  private requireInt(args: ToolArguments, name: string): number {
    const value = args[name];
    if (typeof value !== "number" || !Number.isInteger(value)) {
      throw new Error(`${name} must be an integer`);
    }
    return value;
  }

  private optionalString(args: ToolArguments, name: string): string | undefined {
    const value = args[name];
    return typeof value === "string" && value.trim() ? value : undefined;
  }

  private optionalNumber(args: ToolArguments, name: string): number | undefined {
    const value = args[name];
    return typeof value === "number" && Number.isFinite(value) ? value : undefined;
  }

  private optionalBoolean(args: ToolArguments, name: string): boolean | undefined {
    const value = args[name];
    return typeof value === "boolean" ? value : undefined;
  }

  private optionalInt(args: ToolArguments, name: string): number | undefined {
    const value = args[name];
    return typeof value === "number" && Number.isInteger(value) ? value : undefined;
  }

  private optionalStringArray(args: ToolArguments, name: string): string[] | undefined {
    const value = args[name];
    return Array.isArray(value) && value.every((item) => typeof item === "string")
      ? value as string[]
      : undefined;
  }
}
