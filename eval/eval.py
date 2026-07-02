#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
OfficeEval evaluation tool.

Scores one task or a batch of tasks. Supports both bundled ground-truth answers
and arbitrary agent output directories.

Expected layout:
    OfficeEval/
    - data/level{1,2}/{word,excel,ppt}/{question_id}/
      - input/          # original input files
      - answer/         # reference answer files
      - config.xml      # scoring rubric
    - eval/engines/     # scoring engines (exe + dll)

Usage:
    # Score ground truth answers.
    python eval.py gt                              # all 200 tasks
    python eval.py gt --level 2 --type word        # a level/type subset
    python eval.py gt --question 72153             # one task

    # Score agent outputs.
    python eval.py agent --output-dir outputs/gpt5/

    # Results are saved under results/ by default.
"""

import os
import sys
import json
import shutil
import subprocess
import argparse
import time
from pathlib import Path
from datetime import datetime


# ============================================================
# Path configuration
# ============================================================

REPO_ROOT = Path(__file__).parent.parent.resolve()  # OfficeEval/
DATA_DIR = REPO_ROOT / "data"
ENGINES_DIR = Path(__file__).parent / "engines"
RESULTS_DIR = REPO_ROOT / "results"

# Task type to evaluator executable.
EVALUATORS = {
    "word":  "WordEvaluator.exe",
    "excel": "ExcelEvaluator.exe",
    "ppt":   "PptEvaluator.exe",
}

# Directory names hard-coded by each evaluator executable:
#   WordEvaluator.exe  -> questions/
#   ExcelEvaluator.exe -> excel_questions/
#   PptEvaluator.exe   -> ppt_questions/
ENGINE_QUESTION_DIRS = {
    "word":  "questions",
    "excel": "excel_questions",
    "ppt":   "ppt_questions",
}


# ============================================================
# Helpers
# ============================================================

def get_question_ids(level, qtype):
    """Return sorted question IDs for a level/type pair."""
    qdir = DATA_DIR / f"level{level}" / qtype
    if not qdir.exists():
        return []
    return sorted([d.name for d in qdir.iterdir() if d.is_dir()])


def create_junction(link_path, target_path):
    """Create a Windows directory junction using absolute paths."""
    link = str(Path(link_path).resolve())
    target = str(Path(target_path).resolve())
    subprocess.run(f'rmdir "{link}" 2>nul & mklink /J "{link}" "{target}"',
                   shell=True, capture_output=True, timeout=10)


def remove_junction(link_path):
    """Remove a Windows directory junction."""
    link = str(Path(link_path).resolve())
    subprocess.run(f'rmdir "{link}" 2>nul',
                   shell=True, capture_output=True, timeout=10)


def setup_eval_workspace(level, qtype):
    """
    Prepare the scoring-engine workspace.

    The scoring engines discover tasks through fixed relative directory names.
    This creates a junction under eval/engines/ pointing to the real task data.

    Returns (workspace_path, junctions_to_cleanup).
    """
    workspace = ENGINES_DIR
    data_qdir = DATA_DIR / f"level{level}" / qtype
    engine_dirname = ENGINE_QUESTION_DIRS[qtype]
    junction = workspace / engine_dirname

    # Create junction: engines/{dirname} -> data/level{N}/{type}/.
    create_junction(junction, data_qdir)
    return workspace, [junction]


def cleanup_junctions(junctions):
    """Remove all created junctions."""
    for j in junctions:
        remove_junction(j)


def run_evaluator(workspace, exe_name, question_id):
    """
    Run one scoring executable and return (exit_code, stdout_text).
    The executable is started directly so timeout handling can terminate it.
    """
    exe_path = str((workspace / exe_name).resolve())

    # Stop leftover Office processes before scoring to avoid blocking dialogs.
    subprocess.run(["powershell.exe", "-Command",
        "Get-Process WINWORD,EXCEL,POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force"],
        capture_output=True, timeout=10)

    # Clear Office recovery records that can trigger "Do you still want to open it?" dialogs.
    subprocess.run(["powershell.exe", "-Command",
        "Remove-Item 'HKCU:\\Software\\Microsoft\\Office\\16.0\\Word\\Resiliency' -Recurse -Force -ErrorAction SilentlyContinue;"
        "Remove-Item 'HKCU:\\Software\\Microsoft\\Office\\16.0\\Excel\\Resiliency' -Recurse -Force -ErrorAction SilentlyContinue;"
        "Remove-Item 'HKCU:\\Software\\Microsoft\\Office\\16.0\\PowerPoint\\Resiliency' -Recurse -Force -ErrorAction SilentlyContinue"],
        capture_output=True, timeout=10)

    try:
        result = subprocess.run(
            [exe_path, question_id],
            capture_output=True,
            timeout=60,
            cwd=str(workspace),
        )
        returncode = result.returncode
        stdout_text = decode_evaluator_output(result.stdout)
    except subprocess.TimeoutExpired:
        returncode = 99
        stdout_text = ""
        # A timeout may leave the evaluator and Office process alive.
        subprocess.run(["powershell.exe", "-Command",
            "Get-Process WordEvaluator,ExcelEvaluator,PptEvaluator,WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force"],
            capture_output=True, timeout=10)
    except Exception:
        returncode = -1
        stdout_text = ""

    return returncode, stdout_text


def decode_evaluator_output(stdout_bytes):
    """Decode evaluator stdout, preserving non-ASCII rubric messages when possible."""
    for encoding in ("utf-8-sig", "gbk"):
        try:
            return stdout_bytes.decode(encoding)
        except UnicodeDecodeError:
            continue
    return stdout_bytes.decode("utf-8", errors="replace")


def parse_score(stdout_text):
    """
    Parse the numeric score from evaluator stdout.

    Expected lines include "Score: 18.3 / 20" or "Score: -1 (evaluation error)".
    Returns float or None.
    """
    import re
    for line in stdout_text.split('\n'):
        if 'Score:' in line:
            after = line.split('Score:', 1)[1].strip()
            m = re.match(r'(-?[\d.]+)', after)
            if m:
                return float(m.group(1))
    return None


def parse_error_details(stdout_text):
    """Extract raw failed-criterion messages from evaluator stdout."""
    errors = []
    in_error = False
    for line in stdout_text.split('\n'):
        if 'Error messages:' in line:
            in_error = True
            continue
        if in_error:
            if 'Scoring details:' in line or '==========' in line:
                break
            line = line.strip()
            if line:
                errors.append(line)
    return errors


# ============================================================
# Evaluation core
# ============================================================

def evaluate_single(level, qtype, question_id, source="gt", include_error_details=False):
    """
    Evaluate one task.

    Args:
        level: 1 or 2
        qtype: "word", "excel", "ppt"
        question_id: question ID, such as "72153"
        source: "gt" for bundled answer files, otherwise an output directory
        include_error_details: include raw evaluator failure messages in JSON

    Returns a dict with score, errors, and evaluator status.
    """
    q_dir = DATA_DIR / f"level{level}" / qtype / question_id
    input_dir = q_dir / "input"
    answer_dir = q_dir / "answer"

    if not q_dir.exists():
        return {"question_id": question_id, "type": qtype, "level": level,
                "score": None, "error": "question dir not found"}

    # Choose the files to score.
    if source == "gt":
        source_dir = answer_dir
    else:
        # Agent mode: outputs usually live under {output_dir}/level{N}/{type}/{qid}/input/.
        candidate = Path(source) / f"level{level}" / qtype / question_id / "input"
        if candidate.exists():
            source_dir = candidate
        else:
            # Fallback: files may be placed directly under {output_dir}/level{N}/{type}/{qid}/.
            source_dir = Path(source) / f"level{level}" / qtype / question_id

    if not source_dir.exists():
        return {"question_id": question_id, "type": qtype, "level": level,
                "score": None, "error": f"source dir not found: {source_dir}"}

    # Create a short temporary scoring tree under eval/engines/.
    tmp_root = ENGINES_DIR / "_tmp_eval"
    tmp_q_dir = tmp_root / question_id

    try:
        # Step 1: Create temporary task structure.
        if tmp_q_dir.exists():
            shutil.rmtree(tmp_q_dir)
        tmp_root.mkdir(exist_ok=True)
        tmp_q_dir.mkdir()
        shutil.copy2(q_dir / "config.xml", tmp_q_dir / "config.xml")

        # Step 2: Copy original input files.
        tmp_input = tmp_q_dir / "input"
        shutil.copytree(input_dir, tmp_input)

        # Step 3: Overlay the files being scored.
        for f in source_dir.iterdir():
            if f.is_file():
                shutil.copy2(f, tmp_input / f.name)

        # Step 4: Ensure the conventional filename expected by the evaluator exists.
        expected_names = {
            "word":  ("Word.docx",  ".docx"),
            "excel": ("Excel.xlsx", ".xlsx"),
            "ppt":   ("PPT.pptx",   ".pptx"),
        }
        expected_name, expected_ext = expected_names[qtype]
        expected_file = tmp_input / expected_name
        if not expected_file.exists():
            for f in tmp_input.iterdir():
                if f.suffix == expected_ext:
                    shutil.copy2(f, expected_file)
                    break

        # Step 5: Point the evaluator junction at the temporary tree.
        engine_dirname = ENGINE_QUESTION_DIRS[qtype]
        junction = ENGINES_DIR / engine_dirname
        create_junction(junction, tmp_root)
        exe_name = EVALUATORS[qtype]
        retcode, stdout_text = run_evaluator(ENGINES_DIR, exe_name, question_id)

        # Step 6: Parse evaluator output.
        score = parse_score(stdout_text)
        if score is not None and score < 0:
            score = 0.0
        raw_errors = parse_error_details(stdout_text)

        return {
            "question_id": question_id,
            "type": qtype,
            "level": level,
            "score": score,
            "max_score": 20.0,
            "score_rate": round(score / 20.0 * 100, 1) if score is not None else None,
            "retcode": retcode,
            "num_errors": len(raw_errors),
            "errors": raw_errors if include_error_details else [],
        }

    except Exception as e:
        return {"question_id": question_id, "type": qtype, "level": level,
                "score": None, "error": str(e)}

    finally:
        # Remove the evaluator junction.
        engine_dirname = ENGINE_QUESTION_DIRS[qtype]
        junction = ENGINES_DIR / engine_dirname
        remove_junction(junction)

        # Remove temporary files. Locked Office files can be ignored.
        try:
            shutil.rmtree(tmp_root)
        except Exception:
            pass


def evaluate_batch(levels, qtypes, question_ids=None, source="gt", output_file=None,
                   include_error_details=False):
    """
    Evaluate a batch of tasks.

    Args:
        levels: [1], [2], or [1, 2]
        qtypes: a subset of ["word", "excel", "ppt"]
        question_ids: None for all selected tasks, or a list of question IDs
        source: "gt" or an output directory
        output_file: path for the JSON result file
        include_error_details: include raw evaluator failure messages in JSON
    """
    # Collect tasks.
    tasks = []
    for level in levels:
        for qtype in qtypes:
            qids = get_question_ids(level, qtype)
            if question_ids:
                qids = [q for q in qids if q in question_ids]
            for qid in qids:
                tasks.append((level, qtype, qid))

    total = len(tasks)
    print(f"{'=' * 60}")
    print(f"OfficeEval Evaluation")
    print(f"{'=' * 60}")
    print(f"Mode: {source}")
    print(f"Tasks: {total}")
    print()

    results = []
    start_time = time.time()

    for i, (level, qtype, qid) in enumerate(tasks, 1):
        result = evaluate_single(level, qtype, qid, source, include_error_details)
        results.append(result)

        # Progress output.
        score = result.get("score")
        score_str = f"{score:.1f}/20" if score is not None else "FAIL"
        nerr = result.get("num_errors", 0)
        err_str = f" ({nerr} errors)" if nerr else ""
        elapsed = time.time() - start_time
        eta = elapsed / i * (total - i) if i > 0 else 0

        print(f"  [{i}/{total}] L{level} {qtype:5s} {qid}: "
              f"{score_str}{err_str}  (ETA: {eta:.0f}s)")

    # Print summary.
    print(f"\n{'=' * 60}")
    print(f"SUMMARY")
    print(f"{'=' * 60}")

    for level in sorted(set(t[0] for t in tasks)):
        for qtype in sorted(set(t[1] for t in tasks)):
            subset = [r for r in results if r["level"] == level and r["type"] == qtype]
            if not subset:
                continue
            scores = [r["score"] for r in subset if r["score"] is not None]
            if scores:
                avg = sum(scores) / len(scores)
                perfect = sum(1 for s in scores if s == 20.0)
                print(f"  Level {level} {qtype.upper():5s}: "
                      f"avg={avg:.1f}/20 ({avg/20*100:.1f}%), "
                      f"perfect={perfect}/{len(scores)}, "
                      f"min={min(scores):.1f}, max={max(scores):.1f}")

    all_scores = [r["score"] for r in results if r["score"] is not None]
    if all_scores:
        avg = sum(all_scores) / len(all_scores)
        perfect = sum(1 for s in all_scores if s == 20.0)
        failed = sum(1 for r in results if r["score"] is None)
        print(f"\n  OVERALL: avg={avg:.1f}/20 ({avg/20*100:.1f}%), "
              f"perfect={perfect}/{len(all_scores)}, failed={failed}")

    # Save results.
    if output_file is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = RESULTS_DIR / f"eval_{source}_{timestamp}.json"
    else:
        output_file = Path(output_file)

    output_file.parent.mkdir(parents=True, exist_ok=True)
    output_data = {
        "timestamp": datetime.now().isoformat(),
        "mode": source,
        "total_tasks": len(results),
        "total_scored": len(all_scores),
        "overall_avg_score_rate": round(sum(all_scores) / len(all_scores) / 20 * 100, 1) if all_scores else None,
        "error_details_included": include_error_details,
        "results": results,
    }
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(output_data, f, ensure_ascii=False, indent=2)
    print(f"\nResults saved to: {output_file}")

    return results


# ============================================================
# CLI entry point
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description="OfficeEval evaluation tool",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python eval.py gt                              # score all 200 ground-truth answers
  python eval.py gt --level 2 --type word        # score Level 2 Word only
  python eval.py gt --question 72153             # score one task
  python eval.py agent --output-dir outputs/gpt5 # score model outputs
  python eval.py gt --include-error-details      # include raw rubric messages
        """
    )
    subparsers = parser.add_subparsers(dest="command", help="Evaluation mode")

    # gt subcommand: score bundled ground truth.
    gt_parser = subparsers.add_parser("gt", help="Score bundled ground-truth answer files")
    gt_parser.add_argument("--level", type=int, choices=[1, 2])
    gt_parser.add_argument("--type", choices=["word", "excel", "ppt"])
    gt_parser.add_argument("--question", type=str, help="Question ID to score")
    gt_parser.add_argument("--output", type=str, help="Output JSON path")
    gt_parser.add_argument(
        "--include-error-details",
        action="store_true",
        help="Include raw evaluator failure messages; these may use the original rubric language",
    )

    # agent subcommand: score agent outputs.
    agent_parser = subparsers.add_parser("agent", help="Score files produced by an agent")
    agent_parser.add_argument("--output-dir", required=True, help="Agent output directory")
    agent_parser.add_argument("--level", type=int, choices=[1, 2])
    agent_parser.add_argument("--type", choices=["word", "excel", "ppt"])
    agent_parser.add_argument("--question", type=str, help="Question ID to score")
    agent_parser.add_argument("--output", type=str, help="Output JSON path")
    agent_parser.add_argument(
        "--include-error-details",
        action="store_true",
        help="Include raw evaluator failure messages; these may use the original rubric language",
    )

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        return

    levels = [args.level] if args.level else [1, 2]
    qtypes = [args.type] if args.type else ["word", "excel", "ppt"]
    qids = [args.question] if hasattr(args, 'question') and args.question else None
    source = "gt" if args.command == "gt" else args.output_dir

    evaluate_batch(levels, qtypes, qids, source, args.output, args.include_error_details)


if __name__ == "__main__":
    main()
