# Blood Bank — Dispatch Board

## Product
Flask + SQLite academic prototype for donors, hospital stock, requests and donation records. Operate surface: staff clear requests, donors check eligibility and book. All data is synthetic and marked **Demo data**.

## Visual world
A hospital dispatch desk: cream paper, oxidized-red rail, ink type, amber signals, teal for approved. Paper slips, rubber stamps, split-flap stock counters. No cards, glass, gradients-as-decoration, or neon.

## Tokens
| Role | Value |
|---|---|
| Paper | `#F3EDE1` |
| Raised paper | `#FAF6EC` |
| Sunk paper | `#E9E0CF` |
| Ink | `#1F2421` |
| Ink soft | `#55594F` |
| Ink faint | `#62655A` |
| Oxidized red | `#8C1D18` |
| Deep red / rail | `#6B120F` / `#7E1A16` |
| Amber | `#D9A441` |
| Teal | `#2F5D57` |
| Rule | `#CFC4B2` / `#A99C86` |

Type: **Archivo** for UI (650 headings), **JetBrains Mono** for numerals, refs, labels. Labels are 9–11px uppercase with `0.1–0.16em` tracking; body 13–15px. Headings use tight `-0.02em` tracking.

## Layout
- Desktop: 248px fixed rail, sticky; topbar with live clock and demo stamp; content max 1280px.
- ≤980px: rail becomes top header, one-row horizontally scrollable nav with edge fade and active-tab auto-centering.
- ≤1180px: two-column board collapses to one; stock grid 2-up.
- ≤620px: stock 2-up, forms single column, buttons ≥40px, filters ≥38px.

## Motion
- Stock counters settle with a 260ms split-flap tick; counts ease over 720ms.
- Request approval presses a rotate/scale rubber stamp (`stamp`, 620ms) then lifts the slip away.
- Sections reveal on scroll with 60ms stagger; bars fill from `--w`.
- All motion collapses to instant under `prefers-reduced-motion`.

## Components
`rail`, `topbar`, `stat`, `stock strip`, `paper slip`, `stamp`, `status chip`, `filter chip`, `data table`, `toast`, `record strip`, `timeline`, `auth split`. One radius language: 1–2px hard edges; pills only for status chips and avatars.

## Copy
Plain, operational, no exclamation marks. Numbers are units, refs are `REQ-0000`, dates ISO. Buttons are verbs: Approve, Decline, Complete, Log donation.
