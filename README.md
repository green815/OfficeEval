# OfficeEval Community Reproduction

This repository is an independent community re-implementation and data release
of the OfficeEval benchmark, reconstructed from the benchmark description and
results reported in
["Mind the Gap: Can Frontier LLMs Pass a Standardized Office Proficiency
Exam?"](https://arxiv.org/abs/2606.10956). It is intended to make the Chinese
NCRE Office tasks and scoring workflow easier to inspect and rerun, and is not
affiliated with the original authors.

We appreciate the paper and found the benchmark valuable. At the time this
repository was prepared, we could not find an official public release, so we
put together this community re-implementation for research and reproducibility.

The original paper evaluates LLM agents on 200 practical Microsoft Office
tasks from China's National Computer Rank Examination (NCRE), covering Word,
Excel, and PowerPoint. This release includes the task files, a wrapper around
the bundled machine-grading engines, and reference evaluation outputs.

As a reproduction sanity check, the scores from this release are close to the
numbers reported in the paper:

| Check | Paper | This Release | Notes |
|-------|-------|--------------|-------|
| GPT-5.5 single-turn | `36.1` SR | `37.9` SR | One representative run in `results/eval_gpt-5.5.json` |

These results can be used as reference points after users clone the repo, set
up the evaluator, and run the same evaluation pipeline on their own outputs.

## Dataset

|           | Word | Excel | PPT | Total |
|-----------|------|-------|-----|-------|
| Level 1   | 28   | 28    | 32  | 88    |
| Level 2   | 38   | 37    | 37  | 112   |
| **Total** | **66** | **65** | **69** | **200** |

Each task is scored on a 20-point scale and reported as a 0-100 score rate.
The full benchmark has 7,118 machine-gradable scoring criteria.

## Repository Layout

```text
OfficeEval/
+-- data/                 # 200 Chinese NCRE tasks
+-- eval/
|   +-- eval.py           # scoring entry point
|   +-- engines/          # bundled .NET scoring executables and DLLs
+-- agents/
|   +-- single_turn.py    # single-turn code-generation baseline
|   +-- coding_agent.py   # Claude Code CLI agent runner
|   +-- llm_api.py        # environment-variable API adapter
+-- results/
|   +-- eval_gt.json             # reference score for bundled ground-truth answers
|   +-- eval_gpt-5.5.json        # representative model result
+-- requirements.txt
+-- LICENSE
+-- README.md
```

## Task Format

Each task lives at `data/level{1,2}/{word,excel,ppt}/{question_id}/`:

```text
{question_id}/
+-- input/                 # files the agent should modify
+-- answer/                # reference output, never expose to the agent
+-- config.xml             # scoring rubric, only needed for evaluation
+-- requirement.rtf        # Level 1 task statement
|   or localized .rtf      # Level 2 task statement
+-- task.txt               # cleaned plaintext Chinese instructions
+-- task.json              # structured instructions, including reference images
+-- screenshots/           # rendered input pages/sheets/slides
```

For agent inference, use `task.txt`, `task.json`, original `.rtf`, `input/`,
and optionally `screenshots/`. Do not expose `answer/` or `config.xml`.

## Setup

Evaluation is Windows-only because the scoring engines are .NET Framework
x86 executables and rely on a local Office-compatible environment.

Required for scoring:

- Windows 10/11
- Python 3.8+
- .NET Framework 4.8
- Microsoft Office / Microsoft 365
- **System locale set to Chinese (Simplified, China)**

Required for running agents:

- Python packages in `requirements.txt`
- A configured LLM API for `single_turn.py`, or Claude Code CLI for
  `coding_agent.py`

Install Python dependencies:

```bash
pip install -r requirements.txt
```

## Verify The Evaluator

The release includes `results/eval_gt.json`, produced by scoring the bundled
ground-truth answers. Re-run:

```bash
python eval/eval.py gt --output results/eval_gt_check.json
```

Expected high-level result:

- `total_tasks`: 200
- `total_scored`: 200
- `overall_avg_score_rate`: 95.5

Small differences usually indicate a locale, .NET, path, or Office file
association issue. In our local checks, the drift was concentrated in Word
tasks when the Office/Word environment differed.

The release also includes `results/eval_gpt-5.5.json` as a representative
model result. Its high-level score is `37.9` over 200 scored tasks, compared
with the paper's reported GPT-5.5 average SR of `36.1`. This file is an
evaluation summary for comparison; raw GPT-5.5 output files are not included
in this release. To reproduce the number end to end, run your GPT-5.5 outputs
through the same command used for any agent output:

```bash
python eval/eval.py agent --output-dir results/gpt-5.5 --output results/eval_gpt-5.5_check.json
```

## Evaluate Your Outputs

Your output directory must mirror the data tree. Modified files go under
`input/` and should keep the original filenames.

```text
your_output_dir/
+-- level1/
    +-- excel/
        +-- 68617/
            +-- input/
                +-- excel.xlsx
```

Score all tasks:

```bash
python eval/eval.py agent --output-dir your_output_dir --output results/your_eval.json
```

Score a subset:

```bash
python eval/eval.py agent --output-dir your_output_dir --level 2 --type word
python eval/eval.py agent --output-dir your_output_dir --question 72153
```

The JSON output contains per-task `score`, `score_rate`, evaluator return
code, and failed criterion counts. Raw evaluator failure messages are omitted
by default because they come from the original rubric language. Add
`--include-error-details` if you need those raw messages for debugging.

## Single-Turn Baseline

`agents/single_turn.py` sends `task.json`, screenshots, and input-file context
to one LLM call, extracts Python code, executes it in a copied work directory,
and writes outputs under `results/{model}` by default.

OpenAI-compatible Chat Completions example:

```bash
set OPENAI_API_KEY=...
set OPENAI_BASE_URL=https://api.openai.com/v1
python agents/single_turn.py --model gpt-4.1 --question 68617
```

OpenAI Responses-compatible endpoint:

```bash
set OFFICEEVAL_API_KEY=...
set OFFICEEVAL_BASE_URL=http://host:port/v1
set OFFICEEVAL_API_FORMAT=responses
python agents/single_turn.py --model your-model --level 1 --type excel
```

Anthropic example:

```bash
set OFFICEEVAL_PROVIDER=anthropic
set ANTHROPIC_API_KEY=...
python agents/single_turn.py --model claude-opus-4-1 --question 68617
```

Common options:

```bash
python agents/single_turn.py --model your-model --concurrency 8
python agents/single_turn.py --model your-model --output-dir results/my_run
```

## Coding-Agent Runner

`agents/coding_agent.py` runs Claude Code in an isolated temporary directory
for each task. It copies only the raw exam paper and `input/`, excluding
answers, scoring rubrics, cleaned task text, JSON, and screenshots.

```bash
python agents/coding_agent.py --question 68617
python agents/coding_agent.py --level 2 --type ppt --timeout 3600
python agents/coding_agent.py --output-dir results/my_coding_agent --model claude-opus-4.6
```

If `claude` is not on `PATH`, pass the executable:

```bash
python agents/coding_agent.py --claude-cmd C:\path\to\claude.cmd --question 68617
```

This runner is intended for Windows machines with Office installed. Full
benchmark runs are long-running; use `--question`, `--level`, and `--type` for
smoke tests.

## License

This community reproduction is released under the Creative Commons
Attribution-NonCommercial 4.0 International License (CC BY-NC 4.0), unless
otherwise noted. It is intended for non-commercial research and reproducibility
use only. Commercial use is not permitted. See `LICENSE`.

## Citation

```bibtex
@misc{lv2026mindgapoffice,
  title={Mind the Gap: Can Frontier LLMs Pass a Standardized Office Proficiency Exam?},
  author={Tengchao Lv and Dongdong Zhang and Jiayu Ding and Yilin Jia and Yuzhong Zhao and Yupan Huang and Wenshan Wu and Xiangyang Zhou and Shaohan Huang and Nan Yang and Li Dong and Lei Cui and Furu Wei},
  year={2026},
  eprint={2606.10956},
  archivePrefix={arXiv},
  url={https://arxiv.org/abs/2606.10956}
}
```

## Rights and Contact

This repository is released for research and reproducibility purposes. If the
paper authors plan to release an official version, or if any rights holder
believes any content in this repository infringes their rights, please open an
issue. We will review it promptly and can remove this repository if needed.
