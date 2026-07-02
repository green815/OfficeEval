#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
OfficeEval Single-Turn Agent

This runner reads `task.json`, sends one LLM request to generate Python code,
executes that code against a copied input directory, and saves the resulting
files for evaluation.

Flow:
    task.json + input file paths
         ->
    build the prompt
         ->
    call the LLM once
         ->
    extract returned Python code
         ->
    execute the code in a work directory
         ->
    save output files for evaluation

Usage:
    # Run one task
    python single_turn.py --model claude-sonnet --question 72153

    # Run a subset
    python single_turn.py --model claude-sonnet --level 2 --type word

    # Use a custom output directory
    python single_turn.py --model your-model --output-dir results/your-model/
"""

import os
import sys
import json
import shutil
import re
import time
import argparse
import traceback
import subprocess
from pathlib import Path
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading

from llm_api import call_llm


REPO_ROOT = Path(__file__).parent.parent.resolve()
DATA_DIR = REPO_ROOT / "data"
RESULTS_DIR = REPO_ROOT / "results"


# ============================================================
# Prompt construction
# ============================================================

SYSTEM_PROMPT = """You are an expert in Microsoft Office automation. Your task is to write Python code to manipulate Office documents according to the user's requirements.

Rules:
1. Write complete, directly executable Python code based on the task requirements.
2. Use python-docx for Word documents (.docx), openpyxl for Excel (.xlsx), and python-pptx for PowerPoint (.pptx).
3. The document content and task instructions are in Chinese.
4. Output ONLY the Python code, wrapped in ```python and ```. No explanations."""

# Input-file categories by extension.
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".gif", ".bmp"}
TEXT_EXTENSIONS = {".txt", ".csv", ".xml", ".htm", ".html"}
SCREENSHOTTABLE_EXTENSIONS = {".docx", ".xlsx", ".pptx", ".pdf"}


def build_user_content(task_json_path):
    """
    Build the user content sent to the LLM from task.json, screenshots/, and input/.

    Content structure:
        1. Task requirements from task.json plus embedded RTF images.
        2. Input-file previews as screenshots, original images, or text content.
        3. File listing plus code-generation instructions.

    Returns a list whose items are either {"type": "text", "text": "..."} or
    {"type": "image", "path": "..."}.
    """
    with open(task_json_path, 'r', encoding='utf-8') as f:
        task = json.load(f)

    q_dir = task_json_path.parent
    input_dir = q_dir / "input"
    screenshots_dir = q_dir / "screenshots"

    content = []

    # ---- Part 1: Task requirements ----
    task_text = "## Task Requirements\n\n"
    task_text += f"Input file directory: {input_dir}\n\n"
    task_text += "The following instructions are in Chinese. Follow every step precisely.\n\n"

    for block in task["content"]:
        if block["type"] == "text":
            task_text += block["content"] + "\n\n"
        elif block["type"] == "image":
            # Flush accumulated text before adding an image.
            if task_text.strip():
                content.append({"type": "text", "text": task_text})
                task_text = ""
            # Add an image embedded in the RTF statement.
            img_path = q_dir / block["path"]
            if img_path.exists():
                content.append({"type": "image", "path": str(img_path)})
            task_text = "\n"

    if task_text.strip():
        content.append({"type": "text", "text": task_text})

    # ---- Part 2: Input-file previews ----
    all_files = sorted(
        [f for f in input_dir.iterdir() if f.is_file() and not f.name.startswith("~$")]
    ) if input_dir.exists() else []

    if all_files:
        content.append({"type": "text", "text": "\n## Input Files\n"})

    for f in all_files:
        ext = f.suffix.lower()

        if ext in SCREENSHOTTABLE_EXTENSIONS:
            # Office documents and PDFs use pre-rendered screenshots.
            screenshots = sorted(screenshots_dir.glob(f"{f.name}_*.*")) if screenshots_dir.exists() else []
            if screenshots:
                # Describe what each screenshot represents for this file type.
                ext = f.suffix.lower()
                if ext == ".xlsx":
                    desc = "Each screenshot below shows a different sheet:"
                elif ext == ".pptx":
                    desc = "Each screenshot below shows a different slide:"
                else:
                    desc = "Screenshots of this file's current content:"
                content.append({"type": "text", "text": f"\n### {f.name}\n{desc}\n"})
                for s in screenshots:
                    # Derive the page, slide, or sheet label from the screenshot name.
                    label = s.stem
                    if label.startswith(f.name + "_"):
                        label = label[len(f.name) + 1:]
                    # Add a file-type prefix and avoid duplicated page/slide labels.
                    if ext == ".xlsx":
                        label = f"sheet: {label}"
                    elif ext == ".pptx":
                        label = f"slide: {re.sub(r'^slide', '', label)}"
                    elif ext in (".docx", ".pdf"):
                        label = f"page: {re.sub(r'^page', '', label)}"
                    content.append({"type": "text", "text": f"{label}:"})
                    content.append({"type": "image", "path": str(s)})
            else:
                # Empty documents, such as a zero-slide PPT, may have no screenshots.
                content.append({"type": "text", "text": f"\n### {f.name}\n(empty document, no content to show)\n"})

        elif ext in IMAGE_EXTENSIONS:
            # Send image assets directly.
            content.append({"type": "text", "text": f"\n### {f.name}\n"})
            content.append({"type": "image", "path": str(f)})

        elif ext in TEXT_EXTENSIONS:
            # Inline text-file content.
            text_content = f.read_text(encoding='utf-8', errors='replace')
            content.append({"type": "text", "text": f"\n### {f.name}\n```\n{text_content}\n```\n"})

        # Other asset types, such as .thmx, audio/video, and .rtf, appear only in the file list.

    # ---- Part 3: File listing and instructions ----
    file_list = "\n## File Listing\n\nAll files in the input directory:\n"
    for f in all_files:
        file_list += f"- {f.name}\n"

    file_list += "\n## Instructions\n\n"
    file_list += "Write complete Python code to perform ALL the operations described above. Your code should:\n"
    file_list += "1. Read the input file and understand its structure\n"
    file_list += "2. Perform all required modifications as specified\n"
    file_list += "3. Save the result back to the original file\n"
    file_list += "\nOutput ONLY the code, wrapped in ```python and ```."

    content.append({"type": "text", "text": file_list})

    return content


# ============================================================
# Code extraction and execution
# ============================================================

def extract_code(response_text):
    """
    Extract a Python code block from the LLM response.
    Supports both ```python ... ``` and plain ``` ... ``` fences.
    """
    # Prefer an explicit Python code fence.
    pattern = r'```python\s*\n(.*?)```'
    matches = re.findall(pattern, response_text, re.DOTALL)
    if matches:
        return matches[0].strip()

    # Fall back to any fenced code block.
    pattern = r'```\s*\n(.*?)```'
    matches = re.findall(pattern, response_text, re.DOTALL)
    if matches:
        return matches[0].strip()

    # If no fence exists, treat the entire response as code.
    return response_text.strip()


def execute_code(code, work_dir, timeout=120):
    """
    Execute Python code in the given work directory.

    Returns (success, stdout, stderr).
    """
    code_file = work_dir / "_agent_code.py"
    with open(code_file, 'w', encoding='utf-8') as f:
        f.write(code)

    try:
        result = subprocess.run(
            [sys.executable, str(code_file)],
            capture_output=True,
            text=True,
            timeout=timeout,
            cwd=str(work_dir),
            encoding='utf-8',
            errors='replace',
        )
        return result.returncode == 0, result.stdout, result.stderr
    except subprocess.TimeoutExpired:
        return False, "", "Execution timed out"
    except Exception as e:
        return False, "", str(e)


# ============================================================
# Single-task flow
# ============================================================

def run_single_question(model_name, level, qtype, question_id, output_dir=None):
    """
    Run the complete single-turn flow for one task:
    1. Build the prompt.
    2. Call the LLM.
    3. Extract code.
    4. Execute on a copy of the input files.
    5. Return metadata. Scoring is handled separately by eval/eval.py.
    """
    q_dir = DATA_DIR / f"level{level}" / qtype / question_id
    task_json_path = q_dir / "task.json"

    if not task_json_path.exists():
        return {"error": "task.json not found", "question_id": question_id}

    # Prepare the work directory by copying input/ into it.
    if output_dir:
        work_dir = Path(output_dir).resolve() / f"level{level}" / qtype / question_id
    else:
        work_dir = (RESULTS_DIR / "tmp" / f"level{level}" / qtype / question_id).resolve()
    work_dir.mkdir(parents=True, exist_ok=True)

    # Copy input files into the work directory.
    input_src = q_dir / "input"
    input_dst = work_dir / "input"
    if input_dst.exists():
        shutil.rmtree(input_dst)
    shutil.copytree(input_src, input_dst)

    result = {
        "question_id": question_id,
        "level": level,
        "type": qtype,
        "model": model_name,
        "timestamp": datetime.now().isoformat(),
    }

    try:
        # Step 1: Build the prompt.
        user_content = build_user_content(task_json_path)

        # Step 2: Call the LLM.
        t0 = time.time()
        response_text, usage = call_llm(model_name, SYSTEM_PROMPT, user_content)
        api_time = time.time() - t0

        result["api_time"] = round(api_time, 1)
        result["usage"] = usage

        # Save the raw LLM response.
        with open(work_dir / "llm_response.txt", 'w', encoding='utf-8') as f:
            f.write(response_text)

        # Step 3: Extract code.
        code = extract_code(response_text)
        result["code_length"] = len(code)

        # Save the generated code.
        with open(work_dir / "generated_code.py", 'w', encoding='utf-8') as f:
            f.write(code)

        # Step 4: Execute code against the copied input files.
        # Redirect explicit input paths to the work directory.
        exec_code = code.replace(str(q_dir / "input"), str(input_dst))
        # Also support code that uses input/ as a relative path.
        exec_code = f"import os\nos.chdir(r'{work_dir}')\n\n" + exec_code

        t0 = time.time()
        success, stdout, stderr = execute_code(exec_code, work_dir)
        exec_time = time.time() - t0

        result["exec_success"] = success
        result["exec_time"] = round(exec_time, 1)
        if stdout:
            result["exec_stdout"] = stdout[:500]
        if stderr:
            result["exec_stderr"] = stderr[:500]

        # Save the execution log.
        with open(work_dir / "exec_log.txt", 'w', encoding='utf-8') as f:
            f.write(f"Success: {success}\n\nSTDOUT:\n{stdout}\n\nSTDERR:\n{stderr}")

    except Exception as e:
        result["error"] = str(e)
        result["traceback"] = traceback.format_exc()

    # Save per-task metadata.
    with open(work_dir / "result.json", 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    return result


# ============================================================
# Batch runner
# ============================================================

def run_batch(model_name, levels, qtypes, question_ids=None, output_dir=None, concurrency=1):
    """Run the single-turn agent over a task set, optionally with concurrent API calls."""
    # Collect tasks.
    tasks = []
    for level in levels:
        for qtype in qtypes:
            qtype_dir = DATA_DIR / f"level{level}" / qtype
            if not qtype_dir.exists():
                continue
            for q_dir in sorted(qtype_dir.iterdir()):
                if not q_dir.is_dir():
                    continue
                qid = q_dir.name
                if question_ids and qid not in question_ids:
                    continue
                tasks.append((level, qtype, qid))

    # Resume support: skip tasks that already have result.json.
    remaining = []
    skipped = 0
    for level, qtype, qid in tasks:
        work_dir = Path(output_dir or RESULTS_DIR / "tmp") / f"level{level}" / qtype / qid
        if (work_dir / "result.json").exists():
            skipped += 1
        else:
            remaining.append((level, qtype, qid))

    total = len(remaining)
    print(f"{'=' * 60}")
    print(f"OfficeEval Single-Turn Agent")
    print(f"{'=' * 60}")
    print(f"Model: {model_name}")
    print(f"Tasks: {total} ({skipped} skipped, already done)")
    print(f"Concurrency: {concurrency}")
    if output_dir:
        print(f"Output: {output_dir}")
    print()

    if total == 0:
        print("All tasks already completed.")
        return []

    results = []
    completed = 0
    lock = threading.Lock()
    start_time = time.time()

    def process_task(task_info):
        """Process one task. Safe to call from worker threads."""
        level, qtype, qid = task_info
        return run_single_question(model_name, level, qtype, qid, output_dir)

    if concurrency <= 1:
        # Sequential execution.
        for i, (level, qtype, qid) in enumerate(remaining, 1):
            result = process_task((level, qtype, qid))
            results.append(result)

            success = result.get("exec_success", False)
            error = result.get("error", "")
            status = "OK" if success else ("ERR:" + error if error else "EXEC_FAIL")
            elapsed = time.time() - start_time
            eta = elapsed / i * (total - i) if i > 0 else 0
            tokens = result.get("usage", {}).get("input_tokens", 0) + result.get("usage", {}).get("output_tokens", 0)

            print(f"  [{i}/{total}] L{level} {qtype:5s} {qid}: {status}"
                  f"  ({tokens} tok, {eta:.0f}s left)")
    else:
        # Concurrent execution.
        with ThreadPoolExecutor(max_workers=concurrency) as executor:
            future_to_task = {}
            for task_info in remaining:
                future = executor.submit(process_task, task_info)
                future_to_task[future] = task_info

            for future in as_completed(future_to_task):
                task_info = future_to_task[future]
                level, qtype, qid = task_info

                try:
                    result = future.result()
                except Exception as e:
                    result = {"question_id": qid, "level": level, "type": qtype,
                              "error": str(e), "model": model_name}

                with lock:
                    results.append(result)
                    completed += 1

                    success = result.get("exec_success", False)
                    error = result.get("error", "")
                    status = "OK" if success else ("ERR:" + error if error else "EXEC_FAIL")
                    elapsed = time.time() - start_time
                    eta = elapsed / completed * (total - completed) if completed > 0 else 0
                    tokens = result.get("usage", {}).get("input_tokens", 0) + result.get("usage", {}).get("output_tokens", 0)

                    print(f"  [{completed}/{total}] L{level} {qtype:5s} {qid}: {status}"
                          f"  ({tokens} tok, {eta:.0f}s left)")

    # Print run statistics.
    exec_ok = sum(1 for r in results if r.get("exec_success"))
    exec_fail = sum(1 for r in results if not r.get("exec_success") and not r.get("error"))
    api_err = sum(1 for r in results if r.get("error"))
    total_tokens = sum(r.get("usage", {}).get("input_tokens", 0) + r.get("usage", {}).get("output_tokens", 0) for r in results)
    total_time = time.time() - start_time

    print(f"\n{'=' * 60}")
    print(f"DONE in {total_time:.0f}s")
    print(f"  Exec OK: {exec_ok}/{total}")
    print(f"  Exec Fail: {exec_fail}/{total}")
    print(f"  API Error: {api_err}/{total}")
    print(f"  Total tokens: {total_tokens:,}")

    # Save summary metadata.
    summary_file = Path(output_dir or RESULTS_DIR) / f"agent_{model_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    summary_file.parent.mkdir(parents=True, exist_ok=True)
    summary = {
        "model": model_name,
        "timestamp": datetime.now().isoformat(),
        "total_tasks": total + skipped,
        "completed": total,
        "skipped": skipped,
        "exec_success": exec_ok,
        "total_tokens": total_tokens,
        "total_time_seconds": round(total_time, 1),
        "results": results,
    }
    with open(summary_file, 'w', encoding='utf-8') as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    print(f"Results saved to: {summary_file}")

    return results


# ============================================================
# CLI entry point
# ============================================================

def main():
    parser = argparse.ArgumentParser(description="OfficeEval Single-Turn Agent")
    parser.add_argument("--model", required=True, help="Model name, such as gpt-4.1 or claude-opus-4-1")
    parser.add_argument("--level", type=int, choices=[1, 2], help="Task level to run")
    parser.add_argument("--type", choices=["word", "excel", "ppt"], help="Task type to run")
    parser.add_argument("--question", type=str, help="Question ID to run")
    parser.add_argument("--output-dir", type=str, help="Output directory")
    parser.add_argument("--concurrency", type=int, default=1, help="Number of concurrent API calls (default: 1; recommended: 8-16)")
    args = parser.parse_args()

    levels = [args.level] if args.level else [1, 2]
    qtypes = [args.type] if args.type else ["word", "excel", "ppt"]
    qids = [args.question] if args.question else None
    output_dir = args.output_dir or str(RESULTS_DIR / args.model)

    run_batch(args.model, levels, qtypes, qids, output_dir, args.concurrency)


if __name__ == "__main__":
    main()
