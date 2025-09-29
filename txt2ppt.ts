import fs from "fs/promises";
import path from "path";
import process from "process";
import PptxGenJS from "pptxgenjs";

export type LayoutOption =
  | "LAYOUT_16x9"
  | "LAYOUT_4x3"
  | "LAYOUT_WIDE"
  | "LAYOUT_16x10";

export interface BulletItem {
  text: string;
  bulletType: "bullet" | "number";
  indentLevel: number;
}

export type Block =
  | { type: "paragraph"; text: string }
  | { type: "bullets"; items: BulletItem[] }
  | { type: "image"; alt: string; path: string; sizing?: "cover" | "contain" }
  | { type: "code"; text: string; language?: string }
  | { type: "table"; rows: string[][] };

export interface SlideSpec {
  title: string;
  subtitle?: string;
  blocks: Block[];
  notes?: string;
  background?: string;
}

export interface RenderOptions {
  layout: LayoutOption;
  revision: string;
  meta?: {
    title?: string;
    author?: string;
    company?: string;
  };
  defaultBackground?: string;
}

interface CliOptions {
  inPath: string;
  outPath: string;
  layout: LayoutOption;
  title?: string;
  author?: string;
  company?: string;
  background?: string;
}

type RenderResult =
  | { kind: "rendered"; nextCursor: number }
  | { kind: "split"; nextCursor: number; remainder: Block }
  | { kind: "defer" };

interface RenderDimensions {
  x: number;
  y: number;
  width: number;
  availableHeight: number;
  safeBottom: number;
}

interface ParseState {
  slides: SlideSpec[];
  current?: SlideSpec;
  codeBlock?: { language?: string; lines: string[] };
}

const VALID_LAYOUTS = new Set<LayoutOption>([
  "LAYOUT_16x9",
  "LAYOUT_4x3",
  "LAYOUT_WIDE",
  "LAYOUT_16x10"
]);

const SAFE_MARGIN_X = 0.6;
const SAFE_MARGIN_BOTTOM = 0.7;
const TITLE_Y = 0.4;
const SUBTITLE_Y = 1.0;
const BODY_TOP = 1.6;
const LINE_HEIGHT = 0.35;
const BULLET_LINE_HEIGHT = 0.38;
const CODE_LINE_HEIGHT = 0.32;
const MIN_BLOCK_HEIGHT = 0.6;
const BLOCK_GAP = 0.2;

const DEFAULT_FONT = "BIZ UDPGothic";
const CODE_FONT = "Consolas";
const DEFAULT_LANG = "ja-JP";
const DEFAULT_REVISION = "1";
const MAX_IMAGE_HEIGHT = 3.5;
const TABLE_ROW_HEIGHT = 0.45;
const TABLE_BASE_HEIGHT = 0.6;
const EMU_PER_INCH = 914400;
const MAX_INDENT_LEVEL = 3;

const HEADING_REGEX = /^#\s+/;
const SUBTITLE_REGEX = /^##\s+/;
const CODE_FENCE_REGEX = /^```/;
const NOTE_REGEX = /^>note:/i;
const BACKGROUND_REGEX = /^>bg:/i;
const IMAGE_REGEX = /^\s*!\[(.*?)]\((.+)\)\s*$/;
const BULLET_REGEX = /^(\s*)([-*]|\d+\.)\s+(.*)$/;

export function parseSlides(rawText: string): SlideSpec[] {
  const state: ParseState = { slides: [] };
  const normalized = rawText.replace(/\r\n?/g, "\n");
  const lines = normalized.split("\n");

  for (let index = 0; index < lines.length; index += 1) {
    const rawLine = lines[index];
    const trimmedRight = rawLine.trimEnd();
    const trimmed = trimmedRight.trim();

    if (state.codeBlock) {
      if (CODE_FENCE_REGEX.test(trimmedRight)) {
        appendCodeBlock(state);
        continue;
      }
      state.codeBlock.lines.push(rawLine);
      continue;
    }

    if (!trimmed) {
      continue;
    }

    if (CODE_FENCE_REGEX.test(trimmedRight)) {
      const language = trimmedRight.slice(3).trim() || undefined;
      state.codeBlock = { language, lines: [] };
      continue;
    }

    if (SUBTITLE_REGEX.test(trimmed)) {
      const slide = ensureSlide(state);
      slide.subtitle = trimmed.replace(SUBTITLE_REGEX, "").trim();
      continue;
    }

    if (HEADING_REGEX.test(trimmed)) {
      const title = trimmed.replace(HEADING_REGEX, "").trim();
      startNewSlide(state, title);
      continue;
    }

    if (NOTE_REGEX.test(trimmed)) {
      appendNote(state, trimmedRight);
      continue;
    }

    if (BACKGROUND_REGEX.test(trimmed)) {
      updateBackground(state, trimmedRight);
      continue;
    }

    const imageMatch = rawLine.match(IMAGE_REGEX);
    if (imageMatch) {
      const [, altRaw, pathRaw] = imageMatch;
      const [imagePath, sizingToken] = pathRaw.split(/#(cover|contain)$/i);
      const sizing = sizingToken ? (sizingToken.toLowerCase() === "cover" ? "cover" : "contain") : undefined;
      appendBlock(state, {
        type: "image",
        alt: (altRaw || "").trim(),
        path: imagePath.trim(),
        sizing
      });
      continue;
    }

    if (isTableLine(trimmedRight)) {
      const { rows, nextIndex } = collectTable(lines, index);
      appendBlock(state, { type: "table", rows });
      index = nextIndex;
      continue;
    }

    if (BULLET_REGEX.test(rawLine)) {
      const { items, nextIndex } = collectBullets(lines, index);
      appendBlock(state, { type: "bullets", items });
      index = nextIndex;
      continue;
    }

    const { text: paragraph, nextIndex } = collectParagraph(lines, index);
    appendBlock(state, { type: "paragraph", text: paragraph });
    index = nextIndex;
  }

  appendCodeBlock(state);
  finalizeCurrentSlide(state);
  return state.slides;
}

function startNewSlide(state: ParseState, rawTitle?: string): void {
  finalizeCurrentSlide(state);
  const title = rawTitle?.trim();
  state.current = {
    title: title && title.length > 0 ? title : `Slide ${state.slides.length + 1}`,
    blocks: []
  };
  state.codeBlock = undefined;
}

function ensureSlide(state: ParseState, fallbackTitle?: string): SlideSpec {
  if (!state.current) {
    startNewSlide(state, fallbackTitle ?? `Slide ${state.slides.length + 1}`);
  }
  return state.current!;
}

function finalizeCurrentSlide(state: ParseState): void {
  if (state.current) {
    state.slides.push(state.current);
    state.current = undefined;
  }
}

function appendBlock(state: ParseState, block: Block): void {
  const slide = ensureSlide(state);
  slide.blocks.push(block);
}

function appendNote(state: ParseState, line: string): void {
  const slide = ensureSlide(state);
  const noteText = line.replace(NOTE_REGEX, "").trim();
  slide.notes = slide.notes ? `${slide.notes}\n${noteText}` : noteText;
}

function updateBackground(state: ParseState, line: string): void {
  const slide = ensureSlide(state);
  const backgroundPath = line.replace(BACKGROUND_REGEX, "").trim();
  if (backgroundPath) {
    slide.background = backgroundPath;
  }
}

function appendCodeBlock(state: ParseState): void {
  if (!state.codeBlock) {
    return;
  }
  appendBlock(state, {
    type: "code",
    text: state.codeBlock.lines.join("\n"),
    language: state.codeBlock.language
  });
  state.codeBlock = undefined;
}

function collectTable(lines: string[], startIndex: number): { rows: string[][]; nextIndex: number } {
  const rows: string[][] = [];
  let index = startIndex;

  while (index < lines.length) {
    const candidate = lines[index].trim();
    if (!isTableLine(candidate)) {
      break;
    }
    const inner = candidate.slice(1, candidate.length - 1);
    rows.push(inner.split("|").map(cell => cell.trim()));
    index += 1;
  }

  return { rows, nextIndex: index - 1 };
}

function collectBullets(lines: string[], startIndex: number): { items: BulletItem[]; nextIndex: number } {
  const items: BulletItem[] = [];
  let index = startIndex;

  while (index < lines.length) {
    const raw = lines[index] ?? "";
    const match = raw.match(BULLET_REGEX);
    if (!match) {
      break;
    }
    const indentSpaces = match[1].replace(/\t/g, "  ").length;
    const indentLevel = Math.min(Math.floor(indentSpaces / 2), MAX_INDENT_LEVEL);
    const marker = match[2];
    const textContent = match[3].trim();
    items.push({
      text: textContent,
      bulletType: /^\d+\.$/.test(marker) ? "number" : "bullet",
      indentLevel
    });
    index += 1;
  }

  return { items, nextIndex: index - 1 };
}

function collectParagraph(lines: string[], startIndex: number): { text: string; nextIndex: number } {
  const buffer: string[] = [];
  let index = startIndex;

  while (index < lines.length) {
    const raw = lines[index];
    if (!raw || !raw.trim()) {
      break;
    }
    buffer.push(raw.trim());

    const lookahead = lines[index + 1];
    if (!lookahead || !lookahead.trim()) {
      break;
    }
    if (isBlockBoundary(lookahead)) {
      break;
    }
    index += 1;
  }

  return { text: buffer.join("\n"), nextIndex: index };
}

function isBlockBoundary(line: string): boolean {
  const trimmed = line.trim();
  if (!trimmed) {
    return true;
  }
  return (
    CODE_FENCE_REGEX.test(trimmed) ||
    SUBTITLE_REGEX.test(trimmed) ||
    HEADING_REGEX.test(trimmed) ||
    NOTE_REGEX.test(trimmed) ||
    BACKGROUND_REGEX.test(trimmed) ||
    IMAGE_REGEX.test(line) ||
    isTableLine(trimmed) ||
    BULLET_REGEX.test(line)
  );
}

function isTableLine(line: string): boolean {
  const trimmed = line.trim();
  if (!trimmed.startsWith("|") || !trimmed.endsWith("|")) {
    return false;
  }
  if (/^\|[\s:-]+\|$/.test(trimmed)) {
    return false;
  }
  return trimmed.includes("|");
}

export async function renderSlides(
  pptx: PptxGenJS,
  specs: SlideSpec[],
  options: RenderOptions
): Promise<void> {
  pptx.layout = options.layout;
  pptx.theme = {
    headFontFace: DEFAULT_FONT,
    bodyFontFace: DEFAULT_FONT
  };
  pptx.revision = options.revision;
  if (options.meta?.title) {
    pptx.title = options.meta.title;
  }
  if (options.meta?.author) {
    pptx.author = options.meta.author;
  }
  if (options.meta?.company) {
    pptx.company = options.meta.company;
  }

  const { width: safeWidth, bottom: safeBottom } = computeSafeArea(pptx);

  specs.forEach(spec => {
    const queue: Block[] = [...spec.blocks];
    let sequence = 0;
    let firstSlide = true;

    while (queue.length > 0 || firstSlide) {
      const slideTitle = sequence === 0 ? spec.title : `${spec.title} (cont.)`;
      const slide = pptx.addSlide();

      if (spec.background) {
        slide.background = { path: spec.background };
      } else if (options.defaultBackground) {
        slide.background = { path: options.defaultBackground };
      }

      slide.addText(slideTitle, {
        x: SAFE_MARGIN_X,
        y: TITLE_Y,
        w: safeWidth,
        h: 0.8,
        fontFace: DEFAULT_FONT,
        fontSize: 30,
        bold: true,
        lang: DEFAULT_LANG
      });

      if (spec.subtitle) {
        slide.addText(spec.subtitle, {
          x: SAFE_MARGIN_X,
          y: SUBTITLE_Y,
          w: safeWidth,
          h: 0.6,
          fontFace: DEFAULT_FONT,
          fontSize: 20,
          color: "555555",
          lang: DEFAULT_LANG
        });
      }

      let cursor = BODY_TOP;
      let consumed = false;

      while (queue.length > 0) {
        const block = queue[0];
        const dims: RenderDimensions = {
          x: SAFE_MARGIN_X,
          y: cursor,
          width: safeWidth,
          availableHeight: safeBottom - cursor,
          safeBottom
        };

        const result = renderBlock(slide, block, dims);

        if (result.kind === "defer") {
          break;
        }

        queue.shift();
        consumed = true;

        if (result.kind === "split") {
          queue.unshift(result.remainder);
        }

        cursor = result.nextCursor + BLOCK_GAP;
        if (cursor >= safeBottom - BLOCK_GAP) {
          break;
        }
      }

      if (firstSlide && spec.notes) {
        slide.addNotes(spec.notes);
      }

      sequence += 1;
      firstSlide = false;

      if (queue.length === 0) {
        break;
      }

      if (!consumed) {
        if (cursor !== BODY_TOP) {
          continue;
        }

        const block = queue[0];
        const forced = renderBlock(slide, block, {
          x: SAFE_MARGIN_X,
          y: cursor,
          width: safeWidth,
          availableHeight: safeBottom - cursor,
          safeBottom
        });

        if (forced.kind === "rendered") {
          queue.shift();
        } else if (forced.kind === "split") {
          queue.shift();
          queue.unshift(forced.remainder);
        } else {
          break;
        }
      }
    }
  });
}

function computeSafeArea(pptx: PptxGenJS): { width: number; bottom: number } {
  const width =
    pptx.presLayout.width > 1000
      ? pptx.presLayout.width / EMU_PER_INCH
      : pptx.presLayout.width;
  const height =
    pptx.presLayout.height > 1000
      ? pptx.presLayout.height / EMU_PER_INCH
      : pptx.presLayout.height;

  return {
    width: width - SAFE_MARGIN_X * 2,
    bottom: height - SAFE_MARGIN_BOTTOM
  };
}

function renderBlock(
  slide: PptxGenJS.Slide,
  block: Block,
  dims: RenderDimensions
): RenderResult {
  switch (block.type) {
    case "paragraph":
      return renderParagraph(slide, block, dims);
    case "bullets":
      return renderBullets(slide, block, dims);
    case "image":
      return renderImage(slide, block, dims);
    case "code":
      return renderCode(slide, block, dims);
    case "table":
      return renderTable(slide, block, dims);
    default:
      return { kind: "rendered", nextCursor: dims.y };
  }
}

function renderParagraph(
  slide: PptxGenJS.Slide,
  block: Extract<Block, { type: "paragraph" }>,
  dims: RenderDimensions
): RenderResult {
  const lines = block.text.split(/\n/);
  const neededHeight = Math.max(lines.length * LINE_HEIGHT, MIN_BLOCK_HEIGHT);

  if (neededHeight <= dims.availableHeight) {
    slide.addText(block.text, {
      x: dims.x,
      y: dims.y,
      w: dims.width,
      h: Math.max(neededHeight, MIN_BLOCK_HEIGHT),
      fontFace: DEFAULT_FONT,
      fontSize: 20,
      lineSpacing: 28,
      fit: "shrink",
      lang: DEFAULT_LANG
    });
    return { kind: "rendered", nextCursor: dims.y + neededHeight };
  }

  if (dims.availableHeight < MIN_BLOCK_HEIGHT) {
    return { kind: "defer" };
  }

  const maxLines = Math.max(Math.floor(dims.availableHeight / LINE_HEIGHT) - 1, 1);
  const head = lines.slice(0, maxLines).join("\n");
  const tail = lines.slice(maxLines).join("\n");

  slide.addText(head, {
    x: dims.x,
    y: dims.y,
    w: dims.width,
    h: Math.max(dims.availableHeight, MIN_BLOCK_HEIGHT),
    fontFace: DEFAULT_FONT,
    fontSize: 20,
    lineSpacing: 28,
    fit: "shrink",
    lang: DEFAULT_LANG
  });

  return {
    kind: "split",
    nextCursor: dims.y + dims.availableHeight,
    remainder: { type: "paragraph", text: tail.trimStart() }
  };
}

function renderBullets(
  slide: PptxGenJS.Slide,
  block: Extract<Block, { type: "bullets" }>,
  dims: RenderDimensions
): RenderResult {
  if (block.items.length === 0) {
    return { kind: "rendered", nextCursor: dims.y };
  }

  const neededHeight = Math.max(block.items.length * BULLET_LINE_HEIGHT, MIN_BLOCK_HEIGHT);

  if (neededHeight > dims.availableHeight && dims.availableHeight < MIN_BLOCK_HEIGHT) {
    return { kind: "defer" };
  }

  const maxItems =
    neededHeight <= dims.availableHeight
      ? block.items.length
      : Math.max(Math.floor(dims.availableHeight / BULLET_LINE_HEIGHT) - 1, 1);

  const renderItems: PptxGenJS.TextProps[] = block.items.slice(0, maxItems).map(item => ({
    text: item.text,
    options: {
      bullet: item.bulletType === "number" ? { type: "number" } : true,
      indentLevel: item.indentLevel,
      fontFace: DEFAULT_FONT,
      fontSize: 20,
      lineSpacing: 24,
      lang: DEFAULT_LANG,
      breakLine: true
    }
  }));

  slide.addText(renderItems, {
    x: dims.x,
    y: dims.y,
    w: dims.width,
    h: Math.max(Math.min(neededHeight, dims.availableHeight), MIN_BLOCK_HEIGHT),
    lineSpacing: 24,
    margin: 0.1,
    lang: DEFAULT_LANG
  });

  if (maxItems === block.items.length) {
    return { kind: "rendered", nextCursor: dims.y + Math.min(neededHeight, dims.availableHeight) };
  }

  return {
    kind: "split",
    nextCursor: dims.y + dims.availableHeight,
    remainder: { type: "bullets", items: block.items.slice(maxItems) }
  };
}

function renderImage(
  slide: PptxGenJS.Slide,
  block: Extract<Block, { type: "image" }>,
  dims: RenderDimensions
): RenderResult {
  const available = Math.max(dims.availableHeight, 0);
  if (available < MIN_BLOCK_HEIGHT) {
    return { kind: "defer" };
  }

  const height = Math.max(Math.min(MAX_IMAGE_HEIGHT, available), MIN_BLOCK_HEIGHT);
  const sizingType: "cover" | "contain" = block.sizing === "cover" ? "cover" : "contain";

  slide.addImage({
    path: block.path,
    altText: block.alt || undefined,
    x: dims.x,
    y: dims.y,
    w: dims.width,
    h: height,
    sizing: { type: sizingType, w: dims.width, h: height }
  });

  return { kind: "rendered", nextCursor: dims.y + height };
}

function renderCode(
  slide: PptxGenJS.Slide,
  block: Extract<Block, { type: "code" }>,
  dims: RenderDimensions
): RenderResult {
  const lines = block.text.split(/\n/);
  const neededHeight = Math.max(lines.length * CODE_LINE_HEIGHT + 0.2, MIN_BLOCK_HEIGHT);

  if (neededHeight > dims.availableHeight && dims.availableHeight < MIN_BLOCK_HEIGHT) {
    return { kind: "defer" };
  }

  const fits = neededHeight <= dims.availableHeight;
  const lineLimit = fits
    ? lines.length
    : Math.max(Math.floor((dims.availableHeight - 0.2) / CODE_LINE_HEIGHT) - 1, 1);
  const head = lines.slice(0, lineLimit).join("\n");

  slide.addText(head, {
    x: dims.x,
    y: dims.y,
    w: dims.width,
    h: Math.max(Math.min(neededHeight, dims.availableHeight), MIN_BLOCK_HEIGHT),
    fontFace: CODE_FONT,
    fontSize: 16,
    lineSpacing: 20,
    color: "202020",
    fill: { color: "F2F2F2" },
    fit: "shrink",
    lang: DEFAULT_LANG
  });

  if (fits) {
    return { kind: "rendered", nextCursor: dims.y + Math.min(neededHeight, dims.availableHeight) };
  }

  return {
    kind: "split",
    nextCursor: dims.y + dims.availableHeight,
    remainder: {
      type: "code",
      text: lines.slice(lineLimit).join("\n").trimStart(),
      language: block.language
    }
  };
}

function renderTable(
  slide: PptxGenJS.Slide,
  block: Extract<Block, { type: "table" }>,
  dims: RenderDimensions
): RenderResult {
  const neededHeight = Math.max(TABLE_BASE_HEIGHT + block.rows.length * TABLE_ROW_HEIGHT, MIN_BLOCK_HEIGHT);

  if (neededHeight > dims.availableHeight && dims.availableHeight < MIN_BLOCK_HEIGHT) {
    return { kind: "defer" };
  }

  const targetHeight = Math.max(Math.min(neededHeight, dims.availableHeight), MIN_BLOCK_HEIGHT);
  const tableRows = block.rows.map(row => row.map(cell => ({ text: cell })));

  slide.addTable(tableRows, {
    x: dims.x,
    y: dims.y,
    w: dims.width,
    fontFace: DEFAULT_FONT,
    lang: DEFAULT_LANG
  });

  return { kind: "rendered", nextCursor: dims.y + targetHeight };
}

function parseArguments(argv: string[]): CliOptions {
  const args = argv.slice(2);
  const options: Partial<CliOptions> = { layout: "LAYOUT_16x9" };

  for (let i = 0; i < args.length; i += 1) {
    const token = args[i];
    if (!token.startsWith("--")) {
      throw new Error(`Unexpected argument: ${token}`);
    }
    const [flag, valueFromEquals] = token.split("=", 2);
    const flagName = flag.slice(2);
    const value = valueFromEquals ?? args[++i];
    if (!value) {
      throw new Error(`Missing value for --${flagName}`);
    }
    switch (flagName) {
      case "in":
        options.inPath = value;
        break;
      case "out":
        options.outPath = value;
        break;
      case "layout":
        if (!VALID_LAYOUTS.has(value as LayoutOption)) {
          throw new Error(`Unsupported layout: ${value}`);
        }
        options.layout = value as LayoutOption;
        break;
      case "title":
        options.title = value;
        break;
      case "author":
        options.author = value;
        break;
      case "company":
        options.company = value;
        break;
      case "bg":
        options.background = value;
        break;
      default:
        throw new Error(`Unknown option: --${flagName}`);
    }
  }

  if (!options.inPath) {
    throw new Error("--in is required");
  }
  if (!options.outPath) {
    throw new Error("--out is required");
  }
  if (!options.outPath.toLowerCase().endsWith(".pptx")) {
    throw new Error("--out must point to a .pptx file");
  }

  return options as CliOptions;
}

async function validatePaths(opts: CliOptions): Promise<{ inPath: string; outPath: string }> {
  const inPath = path.resolve(opts.inPath);
  const outPath = path.resolve(opts.outPath);

  try {
    const stats = await fs.stat(inPath);
    if (!stats.isFile()) {
      throw new Error(`${inPath} is not a file`);
    }
  } catch (error) {
    throw new Error(`Failed to read input file: ${(error as Error).message}`);
  }

  const outDir = path.dirname(outPath);
  try {
    const dirStats = await fs.stat(outDir);
    if (!dirStats.isDirectory()) {
      throw new Error(`${outDir} is not a directory`);
    }
  } catch (error) {
    throw new Error(`Output directory missing: ${(error as Error).message}`);
  }

  return { inPath, outPath };
}

async function main(): Promise<void> {
  try {
    const cli = parseArguments(process.argv);
    const { inPath, outPath } = await validatePaths(cli);
    const text = await fs.readFile(inPath, "utf8");
    const specs = parseSlides(text);
    if (specs.length === 0) {
      throw new Error("No slides detected in the input file.");
    }
    const pptx = new PptxGenJS();
    await renderSlides(pptx, specs, {
      layout: cli.layout,
      revision: DEFAULT_REVISION,
      meta: {
        title: cli.title,
        author: cli.author,
        company: cli.company
      },
      defaultBackground: cli.background
    });
    await pptx.writeFile({ fileName: outPath });
  } catch (error) {
    console.error((error as Error).message);
    process.exitCode = 1;
  }
}

if (require.main === module) {
  void main();
}
