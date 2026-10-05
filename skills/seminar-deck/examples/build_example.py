"""Reference build in the lab style: every builder once. Copy next to a talk as build_deck.py.

    set PYTHONUTF8=1
    python build_example.py out.pptx
    powershell -NoProfile -File ../scripts/render.ps1 -src out.pptx -outdir png
    python ../scripts/sheet.py png sheet.png all --cols 3 --scale 0.5
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "scripts"))

from deckkit import LabDeck, arxiv_abs  # noqa: E402
from deckkit.lab import asset  # noqa: E402

FIG = asset("network.png")  # stand-in for a paper figure
OUT = sys.argv[1] if len(sys.argv) > 1 else "example.pptx"

d = LabDeck.create()
d.notes = {"SIM-CoT: step-level supervision": "Заметка докладчика на русском, ключ — точный заголовок слайда."}

d.title_slide("From Chain-of-Thought to Latent Reasoning", "The mainstream, looped LMs and the exotics",
              kicker="Reading club", date="01.10.2026", speakers="Speaker One · Speaker Two")
s = d.content("Roadmap")
for i, name in enumerate(["Explicit CoT", "Coconut", "CODI", "KaVa", "Looped LMs", "Latent RL"]):
    d.capsule_card(s, 0.55 + 3.02 * (i % 3), 1.35 + 1.25 * (i // 3), 2.84, 0.9, f"Block {i + 1}", name,
                   fill="gold" if i == 5 else "warm")
d.divider("Block 1 · CoT → Coconut", "Scaling latents", "More latents or more passes?", number=1)
d.paper("SIM-CoT: step-level supervision", "An auxiliary decoder unrolls each latent into its own step",
        ["The decoder is dropped at inference", "GPT-2 over Coconut: +8.2 points", "Keeps growing at 8–16 latents"],
        FIG, "Where the signal sits", "Takeaway", "Scaling the latent budget needs a step-level signal",
        "Wei et al., ICLR 2026 — SIM-CoT: Supervised Implicit Chain-of-Thought", arxiv_abs("2509.20317"))
d.two_col("PCCoT: parallel latents", "Idea", "All latents are updated together in T passes instead of M",
          "What matters", ["T ≥ M gives Coconut exactly", "KaVa runs this regime: M = 24, T = 3"], FIG, "Jacobi",
          "Wu et al., EMNLP 2025 — Parallel Continuous Chain-of-Thought with Jacobi Iteration", arxiv_abs("2506.18582"))
d.stats("Latent RL beats token RL", [("+4.27", "over explicit GRPO", "4 benchmarks"), ("3.31×", "shorter traces", ""),
                                     ("53.3", "pass@64 on AIME25", "GRPO: 30.0")],
        "Note", "The gap widens with difficulty", "Deng et al., 2026 — Latent-GRPO", arxiv_abs("2604.27998"))
d.cards3("A continuous thought has no density", "Three ways to give it one:",
         [("Anchor", "Discrete token", "A sampled token carries log π", "HRPO"),
          ("Noise", "Gaussian on embedding", "Closed-form density", "Soft Tokens, Hard Truths"),
          ("Mixture", "Gumbel mixture", "Noise on the logits", "SofT-GRPO")], "Each changes what is optimized.")
s = d.content("Bars with explicit labels", "Illustrative numbers · provenance line instead of a citation")
d.bar_chart(s, 0.55, 1.2, 5.0, 3.6, ["Config A", "Config B", "Config C"],
            [("metric 1", (27.5, 4.5, 32.5)), ("metric 2", (28.0, 39.5, 30.0))], decimals=1)
d.hbar_chart(s, 5.8, 1.2, 3.8, 3.6, ["Method A", "Method B", "Method C"], [78.5, 72.0, 79.5],
             ["CDBFC6", "854C65", "5B3445"], labels=["78.5", "72.0", "79.5"])
d.measurement("Our check: a capped budget truncates answers", "Thinking budget capped at 1024 tokens",
              [("", "no cap", "capped"), ("Accuracy", "71.0", "64.5"), ("Truncated traces", "1.5%", "9.0%")],
              [5.2, 3.18, 3.18], "Time = tokens × measured step time", "Why", "Truncation eats the saving",
              "Illustrative numbers · model, benchmark and protocol go here")
d.recap("Block takeaways", [("Collapse", "Latents added one by one stick together"),
                            ("Fix", "A step-level signal"), ("Or", "Refine a fixed block in parallel")], "Next: looped LMs.")
d.meme(asset("molecule-node.png"), 3.3, 0.6, 3.4, [(0.0, 0.85, 1.0, 0.12, "a caption with an outline", 22, "white")])
d.closing("Thank you for your attention!", ["Speaker One · Speaker Two"])
print("notes applied:", d.apply_notes())
d.save(OUT)
print("saved", OUT, len(d.prs.slides), "slides")
