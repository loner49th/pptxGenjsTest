import fs from "fs/promises";
import path from "path";
import process from "process";
import PptxGenJS from "pptxgenjs";

const VALID_LAYOUTS = new Set([
  "LAYOUT_16x9",
  "LAYOUT_4x3",
  "LAYOUT_WIDE",
  "LAYOUT_16x10"
]);

const SAFE_MARGIN_X = 0.6;
const SAFE_MARGIN_TOP = 0.7;
const SAFE_MARGIN_BOTTOM = 0.7;
const TITLE_Y = 0.4;
const SUBTITLE_Y = 1.0;
const BODY_TOP = 1.6;
const LINE_HEIGHT = 0.35;
const BULLET_LINE_HEIGHT = 0.38;
const CODE_LINE_HEIGHT = 0.32;
const MIN_BLOCK_HEIGHT = 0.6;
const BLOCK_GAP = 0.2;

const DEFAULT_FONT = "BIZ UDPゴシック";
const CODE_FONT = "Consolas";

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

  
export function parseSlides(rawText: string): SlideSpec[] {
  const text = rawText.replace(/\r\n?/g, "\n");
  const lines = text.split("\n");
  const slides: SlideSpec[] = [];

  let current: SlideSpec | null = null;
  let codeState: { language?: string; lines: string[] } | null = null;

  const pushCurrent = () => {
    if (current) {
      slides.push(current);
    }
  };

  const ensureSlide = (title: string) => {
    pushCurrent();
    current = { title: title.trim() || `Slide ${slides.length + 1}`, blocks: [] };
    codeState = null;
  };

  const appendNote = (line: string) => {
    if (!current) {
      ensureSlide(`Slide ${slides.length + 1}`);
    }
    const noteText = line.replace(/^>note:/i, "").trim();
    if (!current!.notes) {
      current!.notes = noteText;
    } else {
      current!.notes += "\n" + noteText;
    }
  };

  const setBackground = (line: string) => {
    if (!current) {
      ensureSlide(`Slide ${slides.length + 1}`);
    }
    const bgPath = line.replace(/^>bg:/i, "").trim();
    if (bgPath) {
      current!.background = bgPath;
    }
  };

  for (let i = 0; i < lines.length; i += 1) {
    const rawLine = lines[i];
    const line = rawLine.trimEnd();

    if (codeState) {
      if (/^```/.test(line)) {
        if (!current) {
          ensureSlide(`Slide ${slides.length + 1}`);
        }
        current!.blocks.push({
          type: "code",
          text: codeState.lines.join("\n"),
          language: codeState.language
        });
        codeState = null;
      } else {
        codeState.lines.push(rawLine);
      }
      continue;
    }

    if (/^```/.test(line)) {
      const lang = line.slice(3).trim() || undefined;
      codeState = { language: lang, lines: [] };
      continue;
    }

    if (!line.trim()) {
      continue;
    }

    if (/^#\s+/.test(line)) {
      const title = line.replace(/^#\s+/, "").trim();
      ensureSlide(title);
      continue;
    }

    if (/^##\s+/.test(line)) {
      if (!current) {
        ensureSlide(`Slide ${slides.length + 1}`);
      }
      current!.subtitle = line.replace(/^##\s+/, "").trim();
      continue;
    }

    if (/^>note:/i.test(line)) {
      appendNote(line);
      continue;
    }

    if (/^>bg:/i.test(line)) {
      setBackground(line);
      continue;
    }

    const imageMatch = rawLine.match(/^\s*!\[(.*?)]\((.+)\)\s*$/);
    if (imageMatch) {
      if (!current) {
        ensureSlide(`Slide ${slides.length + 1}`);
      }
      const [, altRaw, pathRaw] = imageMatch;
      const [imgPath, sizingToken] = pathRaw.split(/#(cover|contain)$/i);
      const sizing = sizingToken ? (sizingToken.toLowerCase() === "cover" ? "cover" : "contain") : undefined;
      current!.blocks.push({
        type: "image",
        alt: (altRaw || "").trim(),
        path: imgPath.trim(),
        sizing
      });
      continue;
    }

    if (isTableLine(line)) {
      if (!current) {
        ensureSlide(`Slide ${slides.length + 1}`);
      }
      const tableLines: string[] = [line];
      while (i + 1 < lines.length && isTableLine(lines[i + 1])) {
        tableLines.push(lines[i + 1].trimEnd());
        i += 1;
      }
      const rows = tableLines.map((tableLine) => {
        const trimmed = tableLine.trim();
        const inner = trimmed.slice(1, trimmed.length - 1);
        return inner.split("|").map((cell) => cell.trim());
      });
      current!.blocks.push({ type: "table", rows });
      continue;
    }

    const bulletMatch = rawLine.match(/^(\s*)([-*]|\d+\.)\s+(.*)$/);
    if (bulletMatch) {
      if (!current) {
        ensureSlide(`Slide ${slides.length + 1}`);
      }
      const items: BulletItem[] = [];
      let cursor = i;
      while (cursor < lines.length) {
        const bulletLine = lines[cursor];
        const bulletRaw = bulletLine ?? "";
        const match = bulletRaw.match(/^(\s*)([-*]|\d+\.)\s+(.*)$/);
        if (!match) {
          break;
        }
        const indentSpaces = match[1].replace(/\t/g, "  ").length;
        const indentLevel = Math.min(Math.floor(indentSpaces / 2), 3);
        const marker = match[2];
        const bulletType: "bullet" | "number" = /^\d+\.$/.test(marker) ? "number" : "bullet";
        const textContent = match[3].trim();
        items.push({ text: textContent, bulletType, indentLevel });
        cursor += 1;
      }
      current!.blocks.push({ type: "bullets", items });
      i = cursor - 1;
      continue;
    }

    if (!current) {
      ensureSlide(line);
      continue;
    }

    const paragraphLines: string[] = [rawLine.trim()];
    while (i + 1 < lines.length && lines[i + 1].trim() && !/^(#|##|```|\s*!\[|>note:|>bg:)/.test(lines[i + 1].trim())) {
      const nextLine = lines[++i];
      if (isTableLine(nextLine)) {
        i -= 1;
        break;
      }
      const bulletAhead = nextLine.match(/^(\s*)([-*]|\d+\.)\s+(.*)$/);
      if (bulletAhead) {
        i -= 1;
        break;
      }
      paragraphLines.push(nextLine.trim());
    }
    function needSlide(): SlideSpec {
      if (!current) {
        ensureSlide(`Slide ${slides.length + 1}`);
      }
      return current!;
    }

    // 使い方：
    needSlide().blocks.push({ type: "paragraph", text: paragraphLines.join("\n") });
  }

  if (codeState) {
    if (!current) {
      ensureSlide(`Slide ${slides.length + 1}`);
    }
    current!.blocks.push({
      type: "code",
      text: codeState.lines.join("\n"),
      language: codeState.language
    });
  }

  if (current) {
    slides.push(current);
  }

  return slides;
}

function isTableLine(line: string): boolean {
  const t = line.trim();
  if (!t.startsWith("|") || !t.endsWith("|")) return false;
  if (/^\|[\s:-]+\|$/.test(t)) return false; // 区切り行
  return t.includes("|");
}

export async function renderSlides(pptx: PptxGenJS, specs: SlideSpec[], options: RenderOptions): Promise<void> {
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

  // 単位正規化（EMUならインチに変換）
  const EMU_PER_INCH = 914400;
  const presW_in = pptx.presLayout.width  > 1000 ? pptx.presLayout.width  / EMU_PER_INCH : pptx.presLayout.width;
  const presH_in = pptx.presLayout.height > 1000 ? pptx.presLayout.height / EMU_PER_INCH : pptx.presLayout.height;

  // 以降はインチで計算
  const safeWidth  = presW_in - SAFE_MARGIN_X * 2;
  const safeBottom = presH_in - SAFE_MARGIN_BOTTOM;

  specs.forEach((spec) => {
    const blocksQueue: Block[] = [...spec.blocks];
    let part = 0;
    let isFirstSlide = true;

    while (blocksQueue.length > 0 || isFirstSlide) {
      const slideTitle = part === 0 ? spec.title : `${spec.title} (cont.)`;
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
        lang: "ja-JP"
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
          lang: "ja-JP"
        });
      }
      let cursor = BODY_TOP;
      let consumedInThisSlide = false;

      while (blocksQueue.length > 0) {
        const block = blocksQueue[0];
        const availableHeight = safeBottom - cursor;
        const result = renderBlock(slide, block, {
          x: SAFE_MARGIN_X,
          y: cursor,
          width: safeWidth,
          availableHeight,
          safeBottom
        });

        if (result.kind === "defer") {
          break;
        }

        blocksQueue.shift();
        consumedInThisSlide = true;

        if (result.kind === "split") {
          blocksQueue.unshift(result.remainder);
        }

        cursor = result.nextCursor + BLOCK_GAP;
        if (cursor >= safeBottom - BLOCK_GAP) {
          break;
        }
      }

      if (isFirstSlide && spec.notes) {
        slide.addNotes(spec.notes);
      }

      part += 1;
      isFirstSlide = false;

      if (!consumedInThisSlide && blocksQueue.length > 0) {
        const block = blocksQueue.shift()!;
        blocksQueue.unshift(block);
        if (cursor !== BODY_TOP) {
          continue;
        }
        const forcedResult = renderBlock(slide, block, {
          x: SAFE_MARGIN_X,
          y: cursor,
          width: safeWidth,
          availableHeight: safeBottom - cursor,
          safeBottom
        });
        if (forcedResult.kind === "rendered") {
          blocksQueue.shift();
        } else if (forcedResult.kind === "split") {
          blocksQueue.shift();
          blocksQueue.unshift(forcedResult.remainder);
        }
      }

      if (blocksQueue.length === 0) {
        break;
      }
    }
  });
}

function renderBlock(
  slide: PptxGenJS.Slide,
  block: Block,
  dims: { x: number; y: number; width: number; availableHeight: number; safeBottom: number }
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
  dims: { x: number; y: number; width: number; availableHeight: number }
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
      lang: "ja-JP"
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
    lang: "ja-JP"
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
  dims: { x: number; y: number; width: number; availableHeight: number }
): RenderResult {
  if (block.items.length === 0) {
    return { kind: "rendered", nextCursor: dims.y };
  }
  const neededHeight = Math.max(block.items.length * BULLET_LINE_HEIGHT, MIN_BLOCK_HEIGHT);
  if (neededHeight > dims.availableHeight && dims.availableHeight < MIN_BLOCK_HEIGHT) {
    return { kind: "defer" };
  }
  const maxItems = neededHeight <= dims.availableHeight
    ? block.items.length
    : Math.max(Math.floor(dims.availableHeight / BULLET_LINE_HEIGHT) - 1, 1);

  const renderItems: PptxGenJS.TextProps[] = block.items.slice(0, maxItems).map(item => ({
    text: item.text,
    options: {
      bullet: item.bulletType === "number" ? { type: "number" } : true,
      indentLevel: item.indentLevel,
      fontFace: DEFAULT_FONT, fontSize: 20, lineSpacing: 24, lang: "ja-JP",
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
    lang: "ja-JP"
  });

  if (maxItems === block.items.length) {
    return { kind: "rendered", nextCursor: dims.y + Math.min(neededHeight, dims.availableHeight) };
  }

  const remainder: Block = {
    type: "bullets",
    items: block.items.slice(maxItems)
  };
  return {
    kind: "split",
    nextCursor: dims.y + dims.availableHeight,
    remainder
  };
}

function renderImage(
  slide: PptxGenJS.Slide,
  block: Extract<Block, { type: "image" }>,
  dims: { x: number; y: number; width: number; availableHeight: number }
): RenderResult {
  const height = Math.min(3.5, dims.availableHeight);
  if (height < MIN_BLOCK_HEIGHT) {
    return { kind: "defer" };
  }
  const H = Math.max(Math.min(3.5, dims.availableHeight), MIN_BLOCK_HEIGHT);
  const sizingType = (block.sizing ?? "contain") as "contain" | "cover";
 slide.addImage({
    path: block.path,
    altText: block.alt || undefined,
    x: dims.x, y: dims.y, w: dims.width, h: H,
    sizing: { type: sizingType, w: dims.width, h: H }
  });
  return { kind: "rendered", nextCursor: dims.y + Math.max(height, MIN_BLOCK_HEIGHT) };
}

function renderCode(
  slide: PptxGenJS.Slide,
  block: Extract<Block, { type: "code" }>,
  dims: { x: number; y: number; width: number; availableHeight: number }
): RenderResult {
  const lines = block.text.split(/\n/);
  const neededHeight = Math.max(lines.length * CODE_LINE_HEIGHT + 0.2, MIN_BLOCK_HEIGHT);
  if (neededHeight > dims.availableHeight && dims.availableHeight < MIN_BLOCK_HEIGHT) {
    return { kind: "defer" };
  }
  const fits = neededHeight <= dims.availableHeight;
  const lineLimit = fits ? lines.length : Math.max(Math.floor((dims.availableHeight - 0.2) / CODE_LINE_HEIGHT) - 1, 1);
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
    lang: "ja-JP"
  });

  if (fits) {
    return { kind: "rendered", nextCursor: dims.y + Math.min(neededHeight, dims.availableHeight) };
  }

  const remainder: Block = {
    type: "code",
    text: lines.slice(lineLimit).join("\n").trimStart(),
    language: block.language
  };
  return {
    kind: "split",
    nextCursor: dims.y + dims.availableHeight,
    remainder
  };
}

function renderTable(
  slide: PptxGenJS.Slide,
  block: Extract<Block, { type: "table" }>,
  dims: { x: number; y: number; width: number; availableHeight: number }
): RenderResult {
  const rowHeight = 0.45;
  const baseHeight = 0.6;
  const neededHeight = Math.max(baseHeight + block.rows.length * rowHeight, MIN_BLOCK_HEIGHT);
  if (neededHeight > dims.availableHeight && dims.availableHeight < MIN_BLOCK_HEIGHT) {
    return { kind: "defer" };
  }

  const tableRows = block.rows.map(r => r.map(c => ({ text: c })));
  slide.addTable(tableRows, {
  x: dims.x, y: dims.y, w: dims.width,
    fontFace: DEFAULT_FONT, lang: "ja-JP",
  });

  const usedHeight = Math.min(neededHeight, Math.max(dims.availableHeight, MIN_BLOCK_HEIGHT));
  return { kind: "rendered", nextCursor: dims.y + usedHeight };
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

async function main() {
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
      revision: "1",
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

