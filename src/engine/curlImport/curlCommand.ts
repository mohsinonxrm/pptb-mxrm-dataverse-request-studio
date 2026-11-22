// Read a pasted `curl …` command into method, URL, headers and body (#271).
//
// Accepts what people actually paste: a browser's "Copy as cURL" (bash or
// cmd), DRS's own cURL export, and hand-written commands. Nothing is
// expanded or run — `$select` in double quotes stays `$select`, and a body
// read from a file (`-d @file`) is reported, not read.

export interface CurlHeader {
  name: string;
  value: string;
}

export interface CurlCommand {
  method: string;
  url: string;
  headers: CurlHeader[];
  /** The request body, or undefined when the command sends none. */
  body?: string;
  /** Things the command does that aren't imported (credentials, ignored options). */
  notes: string[];
}

export type CurlParse = { ok: true; command: CurlCommand } | { ok: false; error: string };

// Options that take a value. Anything else that starts with `-` is a flag.
const WITH_VALUE = new Set([
  '-X',
  '--request',
  '-H',
  '--header',
  '-d',
  '--data',
  '--data-raw',
  '--data-binary',
  '--data-ascii',
  '--data-urlencode',
  '--json',
  '-u',
  '--user',
  '-b',
  '--cookie',
  '-c',
  '--cookie-jar',
  '-A',
  '--user-agent',
  '-e',
  '--referer',
  '-o',
  '--output',
  '--url',
  '-m',
  '--max-time',
  '--connect-timeout',
  '-x',
  '--proxy',
  '--retry',
  '-w',
  '--write-out',
  '--oauth2-bearer',
  '-E',
  '--cert',
  '--key',
  '--cacert',
  '-T',
  '--upload-file',
  '-F',
  '--form',
]);

const DATA_OPTIONS = new Set([
  '-d',
  '--data',
  '--data-raw',
  '--data-binary',
  '--data-ascii',
  '--data-urlencode',
  '--json',
]);

/** Short flags that can be bundled (`-sSL`). */
const SHORT_FLAGS = new Set('sSLkivgGIfnqN#0123'.split(''));

/**
 * Undo the escaping of Windows cmd ("Copy as cURL (cmd)"): `^` escapes the
 * next character and ends a continued line. Only applied when the text looks
 * like cmd, so a `^` inside a bash-quoted body is left alone.
 */
function fromCmd(text: string): string {
  if (!/\^"|\s\^\r?\n/.test(text)) return text;
  return text.replace(/\^\r?\n/g, ' ').replace(/\^([\s\S])/g, '$1');
}

/** Split a shell command into words: bash quoting, `$'…'`, line continuations. */
export function shellWords(input: string): string[] {
  const text = fromCmd(input)
    // PowerShell's backtick continuation (`curl.exe … `<newline>).
    .replace(/`\r?\n/g, ' ');
  const words: string[] = [];
  let word = '';
  let inWord = false;
  let i = 0;
  const end = () => {
    if (inWord) words.push(word);
    word = '';
    inWord = false;
  };
  while (i < text.length) {
    const c = text[i];
    if (c === '\\' && (text[i + 1] === '\n' || (text[i + 1] === '\r' && text[i + 2] === '\n'))) {
      i += text[i + 1] === '\r' ? 3 : 2;
      continue;
    }
    if (/\s/.test(c)) {
      end();
      i++;
      continue;
    }
    // A comment (DRS's own cURL code starts with some) runs to the end of the line.
    if (!inWord && c === '#') {
      const nl = text.indexOf('\n', i);
      i = nl < 0 ? text.length : nl;
      continue;
    }
    // A pipe, redirect or second command ends the curl command.
    if (!inWord && (c === '|' || c === ';' || c === '>' || c === '&')) break;
    inWord = true;
    if (c === "'") {
      const close = text.indexOf("'", i + 1);
      if (close < 0) throw new Error('A single-quoted value is never closed.');
      word += text.slice(i + 1, close);
      i = close + 1;
    } else if (c === '$' && text[i + 1] === "'") {
      i += 2;
      for (;;) {
        if (i >= text.length) throw new Error("A $'…' value is never closed.");
        const d = text[i];
        if (d === "'") {
          i++;
          break;
        }
        if (d !== '\\') {
          word += d;
          i++;
          continue;
        }
        const e = text[i + 1];
        const simple: Record<string, string> = {
          n: '\n',
          r: '\r',
          t: '\t',
          '\\': '\\',
          "'": "'",
          '"': '"',
          a: '\x07',
          b: '\b',
          f: '\f',
          v: '\v',
          e: '\x1b',
        };
        if (e in simple) {
          word += simple[e];
          i += 2;
        } else if (e === 'x' && /^[0-9a-fA-F]{1,2}/.test(text.slice(i + 2))) {
          const hex = /^[0-9a-fA-F]{1,2}/.exec(text.slice(i + 2))![0];
          word += String.fromCharCode(parseInt(hex, 16));
          i += 2 + hex.length;
        } else if ((e === 'u' || e === 'U') && /^[0-9a-fA-F]{1,8}/.test(text.slice(i + 2))) {
          const hex = new RegExp(`^[0-9a-fA-F]{1,${e === 'u' ? 4 : 8}}`).exec(
            text.slice(i + 2),
          )![0];
          word += String.fromCodePoint(parseInt(hex, 16));
          i += 2 + hex.length;
        } else {
          word += '\\';
          i++;
        }
      }
    } else if (c === '"') {
      i++;
      for (;;) {
        if (i >= text.length) throw new Error('A double-quoted value is never closed.');
        const d = text[i];
        if (d === '"') {
          i++;
          break;
        }
        if (d === '\\' && '"\\$`\n'.includes(text[i + 1])) {
          if (text[i + 1] !== '\n') word += text[i + 1];
          i += 2;
          continue;
        }
        word += d;
        i++;
      }
    } else if (c === '\\') {
      if (i + 1 < text.length) word += text[i + 1];
      i += 2;
    } else {
      word += c;
      i++;
    }
  }
  end();
  return words;
}

/** Parse a pasted curl command. */
export function parseCurl(input: string): CurlParse {
  let words: string[];
  try {
    words = shellWords(input.trim());
  } catch (e) {
    return { ok: false, error: e instanceof Error ? e.message : String(e) };
  }
  if (words[0] === '$') words.shift();
  const program = words.shift() ?? '';
  if (!/^curl(\.exe)?$/i.test(program.replace(/^.*[\\/]/, ''))) {
    return { ok: false, error: 'Paste a command that starts with curl.' };
  }

  let method: string | undefined;
  let url: string | undefined;
  const headers: CurlHeader[] = [];
  const data: string[] = [];
  const notes: string[] = [];
  let get = false;
  let json = false;
  const ignored = new Set<string>();

  for (let i = 0; i < words.length; i++) {
    let w = words[i];
    let value: string | undefined;
    if (w.startsWith('--') && w.includes('=')) {
      const eq = w.indexOf('=');
      value = w.slice(eq + 1);
      w = w.slice(0, eq);
    }
    if (!w.startsWith('-') || w === '-') {
      if (/^curl(\.exe)?$/i.test(w))
        return { ok: false, error: 'Paste one curl command at a time.' };
      if (url) return { ok: false, error: `The command has more than one URL ("${w}").` };
      url = w;
      continue;
    }
    // Bundled short flags (`-sS`), ending at most in one option that takes a
    // value — attached (`-XPOST`, `-sXPOST`) or in the next word (`-sX POST`).
    if (/^-[^-]/.test(w) && w.length > 2) {
      let j = 1;
      while (j < w.length && SHORT_FLAGS.has(w[j])) {
        if (w[j] === 'G') get = true;
        if (w[j] === 'I') return { ok: false, error: 'A HEAD request (-I) has nothing to import.' };
        j++;
      }
      if (j === w.length) continue;
      const opt = `-${w[j]}`;
      if (WITH_VALUE.has(opt)) {
        value = j + 1 < w.length ? w.slice(j + 1) : undefined;
        w = opt;
      }
    }
    if (WITH_VALUE.has(w) && value === undefined) {
      if (i + 1 >= words.length) return { ok: false, error: `${w} needs a value.` };
      value = words[++i];
    }
    switch (w) {
      case '-X':
      case '--request':
        method = value!.toUpperCase();
        break;
      case '-H':
      case '--header': {
        const colon = value!.indexOf(':');
        if (colon <= 0) {
          // `-H 'Name;'` sends an empty header; `-H 'Name:'` removes one. Name
          // only the first word: a mistyped `Authorization Bearer …` holds a token.
          const shown = value!.split(/[\s;:]/)[0];
          notes.push(`Left out the header "${shown}": it has no value.`);
          break;
        }
        const v = value!.slice(colon + 1).trim();
        if (v) headers.push({ name: value!.slice(0, colon).trim(), value: v });
        break;
      }
      case '--url':
        if (url) return { ok: false, error: 'The command has more than one URL.' };
        url = value;
        break;
      case '-u':
      case '--user':
      case '--oauth2-bearer':
        notes.push(`Left out ${w}: ToolBox signs you in.`);
        break;
      case '-b':
      case '--cookie':
        notes.push(`Left out ${w}: ToolBox signs you in, so cookies aren't needed.`);
        break;
      case '-G':
      case '--get':
        get = true;
        break;
      case '-I':
      case '--head':
        return { ok: false, error: 'A HEAD request (-I) has nothing to import.' };
      case '-F':
      case '--form':
      case '-T':
      case '--upload-file':
        return {
          ok: false,
          error: `${w} uploads a file or form, which import doesn't read. Use the file / image / attachment modes instead.`,
        };
      default:
        if (DATA_OPTIONS.has(w)) {
          if (w === '--json') json = true;
          const raw = value!;
          if (raw.startsWith('@') && w !== '--data-raw') {
            return {
              ok: false,
              error: `The body is read from a file (${raw}). Paste the command with the body inline, for example --data-raw '{…}'.`,
            };
          }
          data.push(raw);
        } else if (!value && !w.startsWith('--')) {
          if (!SHORT_FLAGS.has(w.slice(1))) ignored.add(w);
        } else if (
          ![
            '--compressed',
            '--globoff',
            '--silent',
            '--show-error',
            '--location',
            '--insecure',
            '--include',
            '--verbose',
            '--fail',
            '--http1.1',
            '--http2',
            '-A',
            '--user-agent',
            '-e',
            '--referer',
            '-o',
            '--output',
            '-m',
            '--max-time',
            '--connect-timeout',
            '--retry',
            '-w',
            '--write-out',
            '-c',
            '--cookie-jar',
          ].includes(w)
        ) {
          ignored.add(w);
        }
    }
  }

  if (!url) return { ok: false, error: 'The command has no URL.' };
  if (ignored.size > 0) notes.push(`Ignored ${[...ignored].join(', ')}.`);

  let body: string | undefined = data.length ? data.join('&') : undefined;
  if (get && body !== undefined) {
    url += (url.includes('?') ? '&' : '?') + body;
    body = undefined;
  }
  if (json) {
    if (!headers.some((h) => h.name.toLowerCase() === 'content-type'))
      headers.push({ name: 'Content-Type', value: 'application/json' });
    if (!headers.some((h) => h.name.toLowerCase() === 'accept'))
      headers.push({ name: 'Accept', value: 'application/json' });
  }
  const finalMethod = method ?? (get ? 'GET' : body !== undefined ? 'POST' : 'GET');
  return { ok: true, command: { method: finalMethod, url, headers, body, notes } };
}
