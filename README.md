# pptx-translator

A Python command-line tool that translates PowerPoint (`.pptx`) presentations
**while preserving the original formatting** (fonts, sizes, colors,
alignment, bullet lists, tables, grouped shapes, speaker notes, etc.).

## Key features

- **Automatic source-language detection.** You only need to specify the
  target language; the source language is detected from the presentation's
  own text.
- **Format-preserving translation.** Text is rewritten in place, reusing the
  original run/paragraph formatting (font, size, color, bold/italic,
  bullets, alignment).
- **User-configurable translation exceptions.** Instead of hard-coding
  domain-specific rules in the program, you can supply a plain-text file
  with "do not translate literally" rules, with three selectable modes
  (plain pre-translation substitution, or case-sensitive/case-insensitive
  protected replacement — see
  [Translation exceptions file](#translation-exceptions-file) below). For
  example, this lets you make sure "Práctica 1" is always translated as
  "Practical Lesson 1" instead of a literal "Practice 1", and that an
  acronym like "SGA" is never altered by the translator.
- **Smart handling of figures.** Text boxes that are part of a composed
  figure (grouped shapes, diagrams, flow charts, etc.) are detected and
  their internal line breaks are treated as layout artifacts rather than
  paragraph breaks: all their text is joined into a single string before
  translating it. Detection works by:
  - Any text box that is part of a `GroupShape` is always treated as a
    figure label.
  - Ungrouped text boxes are also treated as figures when a heuristic
    detects that they very likely are (few words per line, little total
    text, small box relative to the slide, and/or no sentence-ending
    punctuation at the end of the lines).
  - After translating, the destination text box is switched to automatic
    text-to-shape resizing (`word_wrap` + `auto_size`) **only when the
    translated text has more than one word**. A single-word result (a
    short label, acronym, or number) is left as-is even if it overflows
    the shape's original bounds, since forcing it to shrink would make it
    unreadable and single words rarely need re-wrapping.
- **Table support.** Text inside table cells is translated the same way as
  regular text boxes (or as figure content, if the table is nested inside a
  group).
- **Speaker notes are translated too.** The notes attached to each slide
  are translated paragraph by paragraph, the same way as regular slide
  content, keeping each run's own formatting (bold, italics, etc.).
- **Optional audio removal.** If requested, any embedded audio object found
  on a slide is removed from the output presentation, together with its
  relationship entries and any leftover reference to it in the slide's
  animation/timing tree, so the resulting file is not flagged by PowerPoint
  as needing repair. This is disabled by default and can be enabled via
  `--remove-audio` or `TRANSLATOR_REMOVE_AUDIO=true`.
- **Slide-number / date fields preserved.** Dynamic PowerPoint fields
  (`<a:fld>`, e.g. the slide number placeholder) are left untouched, since
  PowerPoint recalculates them automatically.
- **Provider flexibility.** The project supports multiple local and remote
  backends: the default `local` provider uses
  [Argos Translate](https://www.argosopentech.com/), while `openai` lets
  you route translations through any OpenAI-compatible endpoint. Models
  based on OpenAI-compatible APIs can offer higher-quality translations,
  but they typically run more slowly and may trigger a much larger number
  of API calls for a single PowerPoint because the presentation text is
  usually split into many small fragments.
- **Rate-limit friendly.** Translation requests are deduplicated and cached
  in-memory so repeated text is only translated once; remote providers also
  honor a configurable request delay and exponential backoff.

## Project structure

```
pptx-translator/
├── translate_pptx.py            # Entry-point script
├── requirements.txt              # Python dependencies
├── .env.example                  # Example configuration (copy to .env)
├── exceptions.example.txt        # Example translation exceptions file
├── src/
│   └── pptx_translator/
│       ├── cli.py                 # Command-line interface
│       ├── config.py              # Settings loaded from environment/.env
│       ├── lang.py                # Source language auto-detection
│       ├── exceptions_list.py     # User-configurable translation exceptions
│       ├── figure_heuristics.py   # Heuristics to detect figure text boxes
│       ├── media.py               # Embedded audio detection/removal
│       ├── pptx_processor.py      # Core translation engine
│       └── translators/
│           ├── base.py              # Common retry/cache/throttle logic
│           ├── factory.py           # Provider selection
│           ├── openai.py            # OpenAI-compatible remote provider
│           ├── local_argos.py       # Local/offline provider (Argos Translate)
│           └── __init__.py          # Exported translator classes
└── tests/
    └── test_core.py             # Unit tests (exceptions + figure heuristics)
```

## Requirements

- Python 3.10+ (tested with Python 3.13).
- Windows, Linux or macOS.
- Internet connection only the first time you translate a given language
  pair with the local provider (to download the model).

## Installation

All commands below use Windows PowerShell syntax; adapt them for
Linux/macOS if needed (the equivalent virtual-environment activation script
on Linux/macOS is `source .venv/bin/activate`).

```powershell
# 1. Move into the project folder
cd pptx-translator

# 2. Create and activate a virtual environment (venv)
python -m venv .venv
.\.venv\Scripts\Activate.ps1        # Windows PowerShell
# source .venv/bin/activate         # Linux/macOS

# 3. Install dependencies
pip install -r requirements.txt

# 4. (Optional) copy the example configuration file and edit it
copy .env.example .env
```

> **Note:** the `local` provider depends on `argostranslate`, which in turn
> pulls in `torch`/`ctranslate2` as transitive dependencies. The first
> `pip install -r requirements.txt` may take a few minutes and download a
> few hundred MB.

## Usage

Basic usage (translates `presentation.pptx` from its auto-detected source
language into English, using the default `local` provider):

```powershell
python translate_pptx.py presentation.pptx -t en
```

This creates `presentation_en.pptx` next to the original file.

When an input file already ends with the target-language suffix (for example,
`presentation_en.pptx` when translating to `en`), it is skipped and a
warning is logged instead of being translated again.

The output file name is always computed as `<original_stem>_<target_lang>.pptx`.
**Important:** if a file with that destination name already exists, it will be
overwritten without prompting. This is intentional for convenience, but it
can cause data loss if you do not pay attention to the final filename.

When the input is a directory, the program looks for `.pptx` files and
translates each one in turn. By default it only processes files directly
inside that directory; use `-r`/`--recursive` to include nested folders as
well.

### Provider selection

The application supports two providers:

- `local` (default): the offline Argos Translate engine.
- `openai`: any OpenAI-compatible API endpoint.
  > AI-based translations can offer greater consistency and overall quality, and they often respect the original formatting and code/symbol-heavy syntax better than a purely local backend. However, this advantage comes at a cost: the OpenAI-compatible provider makes far heavier use of computational and network resources, and for a PowerPoint with many fragmented text elements **the request volume can become very high**.
  >
  > **DISCLAIMER:** the OpenAI-compatible backend is experimental and should be used with caution. Unlike the local Argos backend, it is usually noticeably slower because PowerPoint text is typically fragmented into hundreds of tiny units (titles, labels, bullet items, figure captions, notes, etc.), and each of those may require an individual translation request. In a dense presentation, the number of API calls can become very high, so latency and **cost can increase substantially**. This is a consequence of the format-preserving translation pipeline.


You can either configure the provider via `.env` or override it per run:

```powershell
python translate_pptx.py presentation.pptx -t en --provider openai --api-key "$env:OPENAI_API_KEY" --api-base-url "https://api.openai.com/v1" --model "gpt-4o-mini"
```

### Command-line arguments

| Argument | Required | Description |
|---|---|---|
| `input` | Yes | Path to the input `.pptx` file or to a directory containing `.pptx` files. |
| `-t`, `--target` | Yes | Target language code (ISO 639-1, e.g. `en`, `fr`, `de`). |
| `-s`, `--source` | No | Source language code (ISO 639-1). If omitted, it is auto-detected from the presentation. |
| `-e`, `--exceptions` | No | Path to a [translation exceptions file](#translation-exceptions-file). |
| `-r`, `--recursive` | No | When `input` is a directory, also process `.pptx` files in nested subdirectories. Without it, only files directly in that folder are processed. |
| `-p`, `--provider` | No | Translation backend: `local` or `openai`. Defaults to `TRANSLATOR_PROVIDER` or `local`. |
| `--api-key` | No | API key for the OpenAI-compatible provider. If omitted, the value from `TRANSLATOR_API_KEY`/`OPENAI_API_KEY` is used. |
| `--api-base-url` | No | Base URL for the OpenAI-compatible API. |
| `--model` | No | Model name for the OpenAI-compatible provider. |
| `--request-delay` | No | Seconds to wait between requests on remote providers. Defaults to `TRANSLATOR_REQUEST_DELAY`. |
| `--remove-audio` | No | Removes embedded audio objects from the presentation before saving the output. Disabled by default; can also be enabled via `TRANSLATOR_REMOVE_AUDIO=true`. |
| `-v`, `--verbose` | No | Increases logging, overriding `TRANSLATOR_LOG_LEVEL`: `-v` = `INFO`, `-vv` = `DEBUG`. Without flags, uses `TRANSLATOR_LOG_LEVEL` (see [Logging](#logging)), or `WARNING` if unset. |

### Examples

Translate to French, letting the source language be auto-detected:

```powershell
python translate_pptx.py practica1.pptx -t fr
```

Translate every `.pptx` file directly inside a folder:

```powershell
python translate_pptx.py "presentaciones/" -t en
```

Translate all `.pptx` files recursively under a folder tree:

```powershell
python translate_pptx.py "presentaciones/" -t en -r
```

Preview what would happen without writing any output files:

```powershell
python translate_pptx.py "presentaciones/" -t en --dry-run
```

In dry-run mode, the program still evaluates the same files and output names,
logs the same provider and target-language decisions, and warns when a target
file already exists and would be overwritten, but it does not create or modify
any `.pptx` files on disk.

Translate with the local engine and apply a custom exceptions file:

```powershell
python translate_pptx.py practica1.pptx -t en -e "exceptions.example.txt"
```

Translate via an OpenAI-compatible API:

```powershell
python translate_pptx.py practica1.pptx -t en --provider openai --api-key "$env:OPENAI_API_KEY" --api-base-url "https://api.openai.com/v1" --model "gpt-4o-mini"
```

## Logging

The log level can be configured in two ways:

- **`TRANSLATOR_LOG_LEVEL` environment variable** (or in your `.env` file):
  one of `ERROR`, `WARN` (an alias of `WARNING`), `INFO`, or `DEBUG`.
  Defaults to `WARN` if unset or set to an unrecognized value.
- **`-v`/`-vv` command-line flags**: always take precedence over
  `TRANSLATOR_LOG_LEVEL` when given. `-v` forces `INFO`, `-vv` forces
  `DEBUG`.

```powershell
# Only errors, no warnings
$env:TRANSLATOR_LOG_LEVEL = "ERROR"
python translate_pptx.py practica1.pptx -t en

# Verbose progress information (equivalent to -v)
$env:TRANSLATOR_LOG_LEVEL = "INFO"
python translate_pptx.py practica1.pptx -t en

# -v/-vv always win over the environment variable
python translate_pptx.py practica1.pptx -t en -vv
```

## Translation exceptions file

Some fragments of text must not be translated literally — for example, a
fixed label like "Práctica 1" that has a specific, non-literal equivalent
in the target language ("Practical Lesson 1"), rather than a word-for-word
translation ("Practice 1"), or an acronym like "SGA" that must always be
kept unchanged. Instead of hard-coding rules like these in the program,
they are supplied via a plain-text file passed with `-e`/`--exceptions`.

**File format:** one rule per line, in the form:

```
[<|!|~]source expression = destination expression
```

- Blank lines and lines starting with `#` are treated as comments and
  ignored.
- Both sides are trimmed of surrounding whitespace.
- An optional one-character prefix selects **how** the rule is applied:

  | Prefix | Mode | Matching | Behavior |
  |---|---|---|---|
  | `<` | Pre-translation substitution | case-insensitive | The source expression is replaced by the destination expression **before** the text is sent to the translator, and the result is translated normally along with the rest of the sentence. Use this when the destination text is already written in the target language, but you don't mind the translator adjusting it grammatically while translating the surrounding text. |
  | `!` | Strict, case-sensitive | exact case match | The source expression is matched **exactly** (case-sensitive) in the original text, protected so the translator cannot alter it, and the destination expression is guaranteed to appear verbatim in the final output. Use this for acronyms/identifiers that must never change case or spelling (e.g. `SGA`). |
  | `~` | Strict, case-insensitive | ignores case | Same guarantee as `!`, but the source expression is matched ignoring case. |
  | *(none)* | Same as `~` | ignores case | Kept for backward compatibility with exception files written before this feature existed. |

  The "strict" modes (`!`/`~`) work by temporarily replacing every match
  with an internal placeholder token before calling the translator, and
  substituting the correct destination text back in once translation is
  complete — this is what actually guarantees the exact output, since it
  prevents the translation engine from mistranslating or rewording the
  protected term (e.g. turning "SGA" into "USG"). The placeholder tokens
  are plain lowercase words (no digits, underscores, or punctuation) so
  they survive tokenization/detokenization by the underlying translation
  engines without being corrupted.
- The rules are applied in the file's order.

**Numeric wildcard:** `{num}` matches one or more digits.

- On the left-hand side (source expression), `{num}` matches any run of
  digits in the original text.
- On the right-hand side (destination expression), each `{num}` is
  replaced, in order, by the digits captured by the corresponding `{num}`
  on the left-hand side.

Example (see [`exceptions.example.txt`](exceptions.example.txt) and
[`exceptions.txt`](exceptions.txt)):

```
~Practica {num} = Practical Lesson {num}
~Práctica {num} = Practical Lesson {num}
!SGA = SGA
SGBD = RDBM
```

This turns "Práctica 1" into "Practical Lesson 1" wherever it appears in
the presentation (in any slide, in any text element), while guaranteeing
"SGA" is never altered.

If no `-e`/`--exceptions` file is supplied, no exception rules are applied
and every piece of text is translated as-is.

## Configuration

The project exposes a small set of environment variables and CLI options.
The provider can be selected with `TRANSLATOR_PROVIDER` or `--provider`,
while the remaining settings are intentionally minimal and limited to
runtime behavior such as logging and optional preprocessing.

```dotenv
TRANSLATOR_PROVIDER=local
TRANSLATOR_LOG_LEVEL=WARN
TRANSLATOR_REMOVE_AUDIO=false
```

## How the translation process works internally

1. Load the presentation with `python-pptx`.
2. Optionally remove embedded audio shapes from slides when `--remove-audio`
   or `TRANSLATOR_REMOVE_AUDIO=true` is enabled.
3. If no `--source` was given, auto-detect the source language from an
   aggregated sample of the presentation's text (using `langdetect`).
4. Walk every slide (recursively descending into grouped shapes), every
   table cell, and every slide's speaker notes, collecting translation
   units:
   - **Regular content** (not part of a figure): each paragraph is
     collected and translated independently, preserving bullet/paragraph
     structure and each run's own formatting.
   - **Figure content**: all paragraphs' text is joined into a single
     string (manual line breaks become simple spaces) and translated as
     one unit.
   - **Speaker notes**: treated the same way as regular content (never as
     a figure), paragraph by paragraph.
5. Apply the user-supplied translation exceptions (if any) to every
   collected text before translation.
6. Deduplicate identical texts and translate each unique string only once
   (with caching and retries), to minimize the number of translation calls.
7. Write the translations back:
   - Regular paragraphs (including speaker notes) are rewritten in place,
     keeping their original run formatting and paragraph properties
     (bullets, alignment, indentation).
   - Figure text boxes are collapsed into a single paragraph holding the
     translated (joined) text. Their text frame is switched to automatic
     text-to-shape sizing (`word_wrap=True` +
     `auto_size=MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE`) **only if the translated
     text has more than one word**; single-word results are left
     untouched, even if they overflow the shape.
8. Save the result to the output path.

## Testing

Unit tests cover the translation-exceptions parser/matcher and the
figure-detection heuristics:

```powershell
python -m unittest discover -s tests -v
```

To manually validate the full pipeline end-to-end, use the provided sample
presentation:

```powershell
python translate_pptx.py "P1 Introduction to Oracle\practica1.pptx" -t en -e exceptions.txt -v
```

For maximum detail, use `-vv` to switch to `DEBUG` logging. By default the
tool logs only warnings and errors (`WARN` level, configurable via
`TRANSLATOR_LOG_LEVEL`; see [Logging](#logging)).
## Troubleshooting

- **"Argos Translate does not offer a direct model for 'X' -> 'Y'"**: not
  every language pair has a directly-trained Argos Translate model.
  Translate through English as an intermediate language (translate to
  English first, then translate the English output to your final target
  language) for that pair.
- **First run of the `local` provider is slow**: this is expected, since
  the language model (and, on first use overall, PyTorch/CTranslate2) is
  being downloaded and cached. Subsequent runs with the same language pair
  are fast, since no network calls are made at all.
- **"Invalid exception rule at ..."**: the exceptions file has a line that
  is neither blank, a comment (`#...`), nor of the form
  `[<|!|~]source = destination`. Fix or remove that line.
- **PowerPoint reports the file needs repair after translating a
  presentation with audio**: this was a bug in earlier versions caused by
  leftover references to removed audio shapes in the slide's animation
  ("timing") tree; it has been fixed by also pruning those dangling
  references when audio is removed. If you still see this with a specific
  file, please report it together with the presentation, as there may be
  other kinds of dangling references this tool doesn't yet know to clean up.
- **An exception rule's placeholder token (e.g. something looking like
  `zqkpptxavxq`) leaks into the translated output**: this means the
  translation engine mangled the internal protection token badly enough
  that it could not be matched back, which should be very rare given the
  purely alphabetic scheme used. If it happens, please report the
  offending text/language pair; as a workaround, try the `<` (pre-translation)
  mode for that specific rule instead of `!`/`~`.
