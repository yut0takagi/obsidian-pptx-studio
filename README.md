# PPTX Viewer for Obsidian

View, edit and save PowerPoint decks without leaving Obsidian.

Obsidian natively opens Markdown, images, audio, video, PDF and canvas — `.pptx`
is not on that list, so a deck in your vault is invisible in the file explorer
and cannot be linked to. This plugin registers the extension and renders slides
directly from the OOXML, with no external converter and nothing to install
alongside it.

## Features

- **Deck view** — click a `.pptx` in the file explorer to open it. Thumbnail
  rail, `←`/`→` (or `j`/`k`) to page, zoom, speaker notes.
- **Edit text in place** — double-click a text box, type, `Cmd/Ctrl+S` to save.
  Only the slide XML you touched is rewritten; every other part of the file is
  repacked byte for byte, so animations, transitions, embedded fonts and any
  formatting this plugin does not model survive untouched.
- **Edit like PowerPoint** — a tabbed ribbon (Home, Insert, Format, Design, View,
  and a Table tab that appears when one is selected) over a slide canvas with
  multi-select, marquee selection, drag, resize, rotate, snapping guides,
  clipboard, z-order, align and distribute, group and ungroup, and a right-click
  menu.
- **Format** — font and size, bold/italic/underline/strikethrough, text colour,
  alignment, bullets and numbering, list level, line spacing, vertical anchor,
  shape fill and outline. With text selected the change applies to exactly those
  characters; with only the shape selected it applies to all of its text.
- **Insert** — text boxes, twelve shape presets, lines, pictures from the vault,
  and tables.
- **Tables** — click a cell to select it, shift-click to extend, then insert and
  delete rows and columns, merge and split cells, and fill them.
- **Slides** — add from any layout, duplicate, delete, and reorder by dragging a
  thumbnail. Set a slide's background colour.
- **Rulers and guides** — drag a guide out of a ruler, drag it to move, drag it
  off the slide to remove it. Shapes snap to guides. They are stored in the deck
  the way PowerPoint stores them, so they appear there too.
- **Crop pictures** — the source shows through dimmed while you drag the crop
  handles, and "reset" restores the whole image.
- **Selection pane** — every shape on the slide in stacking order: click to
  select, rename, hide, restack. For the shapes direct manipulation cannot reach.
- **Also** — hyperlinks, format painter, change a shape's preset, rotate and flip
  by 90°, and numeric position, size and rotation fields.
- **Embed slides in notes** — `![[deck.pptx]]` embeds the deck, `![[deck.pptx#3]]`
  pins a single slide.
- **Extract to Markdown** — turn a deck into an outline note so its content is
  searchable, linkable and shows up in backlinks.
- **Export slides as PNG** — at 1×, 2× or 3×.

## Language

The interface follows Obsidian's own display language, so a Japanese vault gets a
Japanese ribbon with nothing to configure. Settings → PPTX Viewer → Language can
pin it to English or Japanese instead.

## Getting around

The ribbon collapses with the chevron at the right of the tab strip, and
remembers whether it was collapsed. In a narrow pane it drops its labels and
group titles rather than scrolling sideways. Selecting a table brings up the
Table tab automatically and deselecting it goes back to where you were.

A small toolbar follows the selection with the things you reach for constantly —
fill, outline, bold, arrange, duplicate, delete, and crop on a picture. Nothing
lives only there; it is a shortcut to what the ribbon already has.

Clicking in the ribbon never takes focus away from the slide, so the arrow keys
keep working immediately afterwards.

## Keyboard

| | |
|---|---|
| `←` `→` / `j` `k` / `PageUp` `PageDown` | previous / next slide |
| `Home` `End` | first / last slide |
| `+` `-` / `0` | zoom in, out, fit to pane |
| `Cmd`/`Ctrl` + wheel | zoom |
| `N` | toggle speaker notes |
| double-click | edit a text box |
| `Esc` | cancel an edit, or clear the selection |
| `Cmd`/`Ctrl` + `Enter` | finish an edit |
| `Tab` / `Shift`+`Tab` (nothing being edited) | walk through the shapes on the slide |
| `Enter` or `F2` | edit the selected shape's text |
| arrow keys (with a shape selected) | nudge 1px, `Shift` for 10px |
| click / `Shift`+click / drag on empty canvas | select, extend, marquee |
| `Cmd`/`Ctrl` + `A` | select every shape on the slide |
| `Cmd`/`Ctrl` + `C` / `X` / `V` / `D` | copy, cut, paste, duplicate |
| `Delete` | delete the selection |
| `Alt` while dragging | drag a copy |
| `Shift` while dragging | constrain to one axis |
| `Shift` while resizing a corner | keep the aspect ratio |
| `Shift` while rotating | snap to 15° |
| `Cmd`/`Ctrl` while dragging | ignore snapping |
| `Tab` / `Shift`+`Tab` while editing text | change the list level |
| `Cmd`/`Ctrl` + `Z` / `Shift+Z` | undo / redo |
| `Cmd`/`Ctrl` + `S` | save |

## How rendering works

The deck is unzipped in memory and each slide is walked into a small model,
then rendered as absolutely-positioned DOM at the deck's native pixel size and
scaled with a CSS transform — so zooming is free and text stays crisp and
selectable.

The parser covers what real decks actually use: theme colour schemes and the
`lumMod`/`lumOff`/`tint`/`shade` transforms, placeholder inheritance from slide
to layout to master, list styles and auto-numbered bullets, gradients, images
with cropping, tables with merged cells, groups, connectors, and SmartArt (via
the plain-shapes fallback drawing PowerPoint stores alongside it).

**Charts** are drawn as SVG — column, bar, stacked and percent-stacked, line,
area, pie, doughnut and scatter — from the values every chart part caches
alongside its embedded workbook. That cache is what the deck was showing when it
was saved, so no spreadsheet has to be opened to plot it. Series take their
colours from the deck's theme accents unless the chart names its own, and a
missing value stays a gap in the line rather than becoming a zero.

Extracting a deck to Markdown turns each chart into its own data table, since
the numbers behind a chart are what someone would want to search for later.

## Editing and safety

Text and shapes inherited from a layout or master are drawn but not editable —
changing them would silently rewrite every slide built on that template.

Editing a run keeps the runs around it intact, so typing inside a bold word
leaves it bold. Only restructuring a paragraph — merging runs, deleting across
them — collapses it onto the formatting of its first run.

Moving a placeholder that inherited its position writes a new `a:xfrm` onto the
slide, which is what PowerPoint does too.

Undo and redo work on whole package parts rather than on individual elements. An
edit that adds a part — inserting a picture, adding a slide — cannot be expressed
as a change to an element, so recording parts is what lets one mechanism cover
every command and guarantees undo lands on a state the parser has already
accepted. The test suite asserts exactly that: after running all sixteen editor
commands and undoing them, every part is byte-identical to where it started.

The first time you save a deck, a one-off copy is written beside it as
`<name>.pptx.bak`. Binary files do not go through Obsidian's file recovery, so
that copy is your undo of last resort.

If the file changed on disk since you opened it, the save is refused rather than
overwriting someone else's work.

## Why it stays responsive

An edit re-derives only the slide it touched. `DeckEditor` reports which package
parts a command changed, and a change confined to one slide's part and its
relationships rebuilds that slide alone; only a change to the presentation, a
layout or a master rebuilds the deck. On a 27-slide deck that is 1.7ms instead of
61ms, and one slide re-rendering instead of all of them.

Rendered slides are cached per index and only the edited one is dropped.
Thumbnails render lazily as they scroll into view, and refresh individually.
The ribbon, selection pane, floating toolbar and status bar all refresh together
in one animation frame rather than each reacting to every event.

```
                      whole deck    one slide    render one slide
21 slides, 702 shapes    30.0ms        0.58ms          4.5ms
27 slides, 233 shapes    60.6ms        1.67ms          6.1ms
```

`npm run smoke -- --bench` prints these for any deck (set `PPTX_SMOKE_BENCH=1`).

## Development

```bash
npm install
npm run dev      # watch build
npm run build    # typecheck + production bundle
npm run smoke -- ~/Downloads/*.pptx   # parse + save round trip against real decks
```

The smoke test runs the parser *and the renderer* under Node with jsdom, then
drives the same code paths the UI does: it edits a run in the rendered DOM,
deletes a paragraph, moves and resizes shapes, inserts text boxes, shapes,
tables and pictures, groups and ungroups, reorders, aligns, formats, adds,
duplicates, reorders and deletes slides — then undoes the lot and checks the
package came back byte-for-byte.
