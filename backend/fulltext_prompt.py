"""
fulltext_prompt.py — Prompt template for the full-text screening phase.

The full-text prompt differs from the TIAB prompt:
  * the model now has the *complete* article (PDF attached natively), not only
    the title/abstract;
  * the model is asked to evaluate each inclusion/exclusion criterion against
    full-text evidence and reference specific sections/figures when relevant;
  * the rationale must be richer than the TIAB rationale (no 12-word cap), so
    that human reviewers can audit the decision quickly.

Both `render_prompt(...)` and `preview_prompt(...)` produce the exact text that
will be sent to the LLM — `preview_prompt` exists so the frontend can show the
user what is going out before the job runs.
"""

from typing import List


def render_prompt(
    synopsis: str,
    inclusion: List[str],
    exclusion: List[str],
    filename: str = "",
) -> str:
    """Render the full-text screening prompt for one article.

    Parameters
    ----------
    synopsis    : study synopsis / PICO description for the systematic review
    inclusion   : list of inclusion criteria (one per line)
    exclusion   : list of exclusion criteria (one per line)
    filename    : optional file name (used for the model's context only)
    """
    inc_lines = "\n".join(f"- {c.strip()}" for c in inclusion if c and c.strip()) \
        or "- (none provided)"
    exc_lines = "\n".join(f"- {c.strip()}" for c in exclusion if c and c.strip()) \
        or "- (none provided)"

    file_hint = f"File: {filename}\n\n" if filename else ""

    return (
        "You are a knowledgeable AI assistant performing FULL-TEXT screening "
        "for a systematic review. The complete article PDF is attached. Read "
        "the full text — methods, results, figures and tables included — and "
        "decide whether the study meets every inclusion criterion and avoids "
        "every exclusion criterion. Unlike TIAB screening, you should ground "
        "your judgement in concrete evidence from the article, not only in "
        "what is hinted at in the abstract.\n"
        "\n"
        f"{file_hint}"
        f"Synopsis/PICO: {synopsis.strip() or '(not provided)'}\n"
        "\n"
        "Inclusion Criteria:\n"
        f"{inc_lines}\n"
        "\n"
        "Exclusion Criteria:\n"
        f"{exc_lines}\n"
        "\n"
        "Instructions:\n"
        "1. Identify the article's PICO elements (population, intervention/exposure, "
        "comparator, outcomes) and the study design from the full text.\n"
        "2. For each inclusion criterion, judge whether the study fulfills it. "
        "Cite the section that supports your judgement when possible.\n"
        "3. For each exclusion criterion, judge whether the study triggers it.\n"
        "4. Treat unspecified details as 'unclear' rather than failing the criterion.\n"
        "5. Perform the above reasoning internally — do not output the steps.\n"
        "\n"
        "Decision logic (full-text is the definitive screening step):\n"
        "- Include only if every inclusion criterion is met (or clearly met) "
        "AND no exclusion criterion is triggered.\n"
        "- Exclude if any inclusion criterion is clearly unmet OR any "
        "exclusion criterion is clearly triggered.\n"
        "- If any criterion is genuinely unclear after reading the full text, "
        "mark the article as 'maybe' so a human reviewer can re-check.\n"
        "\n"
        "Output (JSON only): Return a single JSON object with keys:\n"
        "- title: extracted study title (string)\n"
        "- decision: \"include\" | \"exclude\" | \"maybe\"\n"
        "- rationale: 2-4 sentences explaining the decision and the criteria it hinges on\n"
        "- inclusion_evaluation: array of "
        "{ \"criterion\": string, \"status\": \"met\"|\"unclear\"|\"unmet\", \"evidence\": string }\n"
        "- exclusion_evaluation: array of "
        "{ \"criterion\": string, \"status\": \"met\"|\"unclear\"|\"unmet\", \"evidence\": string }\n"
        "No other text should be produced outside the JSON.\n"
        "\n"
        "Example format:\n"
        "{\n"
        "  \"title\": \"Effect of X on Y in adults with T2D: a randomized trial\",\n"
        "  \"decision\": \"include\",\n"
        "  \"rationale\": \"RCT in adults with T2D evaluating intervention X vs placebo and reporting HbA1c at 12 weeks — all PICO elements met. No exclusion criterion triggered (humans, not pediatric, not in-vitro).\",\n"
        "  \"inclusion_evaluation\": ["
        "{ \"criterion\": \"population: adults with T2D\", \"status\": \"met\", "
        "\"evidence\": \"Methods §2.1 reports 240 adults with T2D, mean age 58.\" }"
        "],\n"
        "  \"exclusion_evaluation\": ["
        "{ \"criterion\": \"non-human study\", \"status\": \"unmet\", "
        "\"evidence\": \"Clinical trial in humans.\" }"
        "]\n"
        "}\n"
        "\n"
        "Now read the attached full-text PDF and output the JSON decision."
    )


def preview_prompt(
    synopsis: str,
    inclusion: List[str],
    exclusion: List[str],
    filename: str = "<article.pdf>",
) -> str:
    """Return the exact prompt that would be sent for an article with this filename.

    Used by the frontend's 'Preview prompt' button — identical to render_prompt
    but the default `filename` is a placeholder when the user has not chosen
    one yet.
    """
    return render_prompt(synopsis, inclusion, exclusion, filename)
