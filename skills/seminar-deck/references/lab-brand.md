# BrainLab brand reference

Extracted from the official template `Шаблон_презентации_BrainLab.pptx`
(theme name "BRAIN LAB BASE"). These are the ground-truth values the helper uses.

## Identity
**BRAInLab** — Basic Research of Artificial Intelligence Laboratory, MIPT.
The wordmark is Latin ("BRAIn lab") with the English tagline
"basic research of artificial intelligence". Because the logo is already in Latin
script, **the same logo serves Russian and English decks** — there is no separate
English logo. Only body copy changes language.

## Format
- Aspect ratio 16:9, slide size **10″ × 5.625″** (`pres.layout = "LAYOUT_16x9"`).
- Brand font **Trebuchet MS** (ships with Microsoft Office; on Linux the PDF renderer
  substitutes a slightly wider face — leave ~10% slack and QA by eye).

## Palette (hex, no `#`)
| Token | Hex | Role |
|---|---|---|
| `maroon` | `854C65` | primary **dark** background (title/section/closing) |
| `maroonDk` | `5B3445` | darker maroon |
| `maroonDeep` | `412430` | deepest maroon; oversized section numbers, body ink |
| `cream` | `FAF8F1` | primary **light** background (content slides), light cards |
| `creamWarm` | `FBEDD6` | warm cream variant |
| `blue` | `9DE1FC` | light blue — the capsule tabs, accents |
| `sky` | `89DAFB` | sky-blue variant |
| `blueDeep` | `0681B2` | deep blue — capsule outline, strong accent, stat numbers |
| `gold` | `F0C987` | gold — accent cards on cream, highlights |
| `rose` | `C397AA` | muted rose |

Text roles: `onMaroon FAF8F1`, `subMaroon E4D2DB` (muted on maroon),
`ink 412430` (body on cream), `subCream 8A6677` (muted on cream).

## Motifs
1. **Molecular network.** Several uses: a faint background texture (`moleculeWatermark`),
   a bright blue+gold accent node (`molecule "node"`), a cream "B" line-network
   (`molecule "network"`), a bright neon "B" outline (`molecule "glow"`), and a 3D glossy
   "B" cluster (`molecule "cluster"`, the hero graphic). One per slide as an accent, never
   behind live text.
2. **Signature capsule cards.** Rounded rectangles (cream or gold) each with a small blue
   "capsule/pill" tab straddling the top-left corner that carries a short label. This
   pill-on-card is *the* recognizable BrainLab element.
3. **Footer mark.** A small molecule "B" mark sits bottom-left on content/section slides.
4. **Split panels.** A maroon half beside a cream half, used for two parallel topics or
   before/after framing (`splitSlide`).
5. **Molecule hero.** A large 3D cluster / network graphic as a full focal point with a
   short statement (`heroMolecule`).

## Structure ("sandwich")
- **Title** — maroon, light logo top-left, big title + subtitle on the LEFT, molecule node
  accent on the RIGHT (text and graphic are horizontally separated so they never overlap).
- **Section divider** — maroon, oversized two-digit number ("01") upper-left + section
  title lower.
- **Content** — cream, title top-left, footer mark + auto page number; fill with capsule
  cards / stats.
- **Split / Hero** — change-of-pace layouts between content runs (unnumbered).
- **Closing** — maroon, light logo, "Спасибо за внимание" (or EN) + contacts on the left,
  molecule cluster on the right.

## Page numbers
Content slides auto-number 1, 2, 3… in creation order; title, section, split, hero and
closing slides are unnumbered. Page numbers sit bottom-right, the footer mark bottom-left.

## Do / Don't
**Do:** keep one idea per slide; sandwich maroon↔cream; gold cards on cream, cream cards
on maroon; 1–2 word capsule labels; generous margins (~0.5″).
**Don't:** add edge stripes, accent bars, or a rule line under titles; invent colors
outside the palette; stack multiple molecules on one slide; put a molecule behind text.
