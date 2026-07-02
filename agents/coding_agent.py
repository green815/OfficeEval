#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
OfficeEval coding-agent runner using the Claude Code CLI.

For each task:
    1. Create an isolated work directory under the system temp directory.
    2. Copy only the task statement and input files, excluding answer/config.xml.
    3. Run Claude Code CLI so it can inspect files, write code, execute, and debug.
    4. Copy the modified input/ directory and Claude Code output back to results/.
    5. Clean up the temp directory.

Isolation design: work_dir is placed under temp and is fully separated from
the OfficeEval repository, so the agent cannot discover ground truth through
the repository layout.

Usage:
    # Run all tasks
    python coding_agent.py

    # Run a subset
    python coding_agent.py --level 1 --type word

    # Run one task
    python coding_agent.py --question 68616

    # Use a custom output directory
    python coding_agent.py --output-dir results/coding-agent-test
"""

import os
import sys
import json
import shutil
import time
import tempfile
import argparse
import subprocess
import threading
from pathlib import Path
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

try:
    import psutil
    HAS_PSUTIL = True
except ImportError:
    HAS_PSUTIL = False


REPO_ROOT = Path(__file__).parent.parent.resolve()
DATA_DIR = REPO_ROOT / "data"
RESULTS_DIR = REPO_ROOT / "results"

CLAUDE_CMD = os.environ.get("CLAUDE_CMD", "claude")
CLAUDE_MODEL = os.environ.get("CLAUDE_MODEL", "claude-opus-4.6")

TASK_PROMPT = (
    "Read the .rtf file in this directory for the task requirements. "
    "The input files to modify are in ./input/. "
    "Complete all the operations described and save all changes back to the original files."
)

# Files and directories excluded from the agent work directory.
# The runner keeps only the raw RTF task statement and input/, matching what a
# human examinee would receive.
EXCLUDE = {"answer", "config.xml", "README.md", "task.txt", "task.json", "screenshots"}
# task_img_*.png is also excluded by prefix in setup_tmp_dir.

# Idle watchdog: no stdout/stderr activity for this many seconds is treated as a hang.
IDLE_TIMEOUT = 600


def kill_proc_tree(pid, timeout=5):
    """Kill a process and all descendants, with Windows-friendly fallback behavior."""
    if HAS_PSUTIL:
        try:
            parent = psutil.Process(pid)
            children = parent.children(recursive=True)
            for c in children:
                try: c.kill()
                except Exception: pass
            try: parent.kill()
            except Exception: pass
            psutil.wait_procs(children + [parent], timeout=timeout)
            return
        except Exception:
            pass
    # Fall back to taskkill /T for process-tree termination.
    try:
        subprocess.run(
            ["taskkill", "/F", "/T", "/PID", str(pid)],
            capture_output=True, timeout=timeout,
        )
    except Exception:
        pass


def cleanup_orphan_claude_procs():
    """Clean up orphaned Claude CLI processes left by previous runs."""
    if not HAS_PSUTIL:
        return 0
    killed = 0
    my_pid = os.getpid()
    for p in psutil.process_iter(['pid', 'name', 'ppid']):
        try:
            name = (p.info['name'] or '').lower()
            if name not in ('claude.exe', 'claude'):
                continue
            # Skip processes whose parent is still alive and is not this runner.
            ppid = p.info['ppid']
            try:
                parent = psutil.Process(ppid)
                if parent.is_running() and parent.pid != my_pid:
                    continue
            except psutil.NoSuchProcess:
                pass
            kill_proc_tree(p.info['pid'])
            killed += 1
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return killed


def setup_tmp_dir(q_dir):
    """
    Create an isolated temp directory and copy task-visible files into it.
    Ground truth and scoring configuration are excluded. The caller owns cleanup.
    """
    tmp_dir = Path(tempfile.mkdtemp(prefix="officeeval_"))

    for item in q_dir.iterdir():
        if item.name in EXCLUDE:
            continue
        if item.name.startswith("task_img_"):
            continue
        dst = tmp_dir / item.name
        if item.is_dir():
            shutil.copytree(item, dst)
        else:
            shutil.copy2(item, dst)

    return tmp_dir


def run_claude_code(work_dir, timeout=900, idle_timeout=IDLE_TIMEOUT):
    """
    Run Claude Code CLI under work_dir and return the parsed JSON result.
    Uses streaming stdout/stderr plus two timeout guards:
      - total timeout
      - idle timeout, triggered when no output is produced for too long
    On timeout, kill the process tree and raise TimeoutExpired.
    """
    cmd = [
        CLAUDE_CMD,
        "--print",
        "--bare",
        "--model", CLAUDE_MODEL,
        "--dangerously-skip-permissions",
        "--output-format", "json",
        "--no-session-persistence",
        TASK_PROMPT,
    ]

    # CREATE_NEW_PROCESS_GROUP helps taskkill /T terminate the whole group.
    creationflags = 0
    if sys.platform == "win32":
        creationflags = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)

    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        cwd=str(work_dir),
        encoding="utf-8",
        errors="replace",
        creationflags=creationflags,
    )

    stdout_chunks = []
    stderr_chunks = []
    last_output_at = [time.time()]
    done = threading.Event()

    def reader(stream, sink):
        try:
            for line in stream:
                sink.append(line)
                last_output_at[0] = time.time()
        except Exception:
            pass
        finally:
            try: stream.close()
            except Exception: pass

    t_out = threading.Thread(target=reader, args=(proc.stdout, stdout_chunks), daemon=True)
    t_err = threading.Thread(target=reader, args=(proc.stderr, stderr_chunks), daemon=True)
    t_out.start(); t_err.start()

    start = time.time()
    timed_out = False
    timeout_reason = None
    while True:
        rc = proc.poll()
        if rc is not None:
            break
        now = time.time()
        if now - start > timeout:
            timed_out = True
            timeout_reason = f"total_timeout ({timeout}s)"
            break
        if now - last_output_at[0] > idle_timeout:
            timed_out = True
            timeout_reason = f"idle_timeout ({idle_timeout}s no output)"
            break
        time.sleep(1)

    if timed_out:
        kill_proc_tree(proc.pid)
        # Give child processes a short window to exit after tree-kill.
        try: proc.wait(timeout=10)
        except subprocess.TimeoutExpired:
            try: proc.kill()
            except Exception: pass
        t_out.join(timeout=2); t_err.join(timeout=2)
        raise subprocess.TimeoutExpired(cmd, timeout, output=''.join(stdout_chunks), stderr=timeout_reason)

    t_out.join(timeout=5); t_err.join(timeout=5)
    stdout_str = ''.join(stdout_chunks)
    stderr_str = ''.join(stderr_chunks)

    try:
        data = json.loads(stdout_str)
    except json.JSONDecodeError:
        data = {
            "raw_stdout": stdout_str[:2000],
            "raw_stderr": stderr_str[:2000],
            "returncode": proc.returncode,
        }

    return data


def process_question(level, qtype, qid, output_dir, timeout=900, max_retries=5):
    """Process one task: temp isolation, Claude Code, then result collection.
    Refusals are retried with a fresh temp directory, up to max_retries attempts.
    """
    q_dir = DATA_DIR / f"level{level}" / qtype / qid
    save_dir = Path(output_dir) / f"level{level}" / qtype / qid
    save_dir.mkdir(parents=True, exist_ok=True)

    last_result = None
    refusal_history = []

    for attempt in range(1, max_retries + 1):
        result = {
            "question_id": qid,
            "level": level,
            "type": qtype,
            "timestamp": datetime.now().isoformat(),
            "attempt": attempt,
        }

        tmp_dir = None
        try:
            tmp_dir = setup_tmp_dir(q_dir)

            t0 = time.time()
            data = run_claude_code(tmp_dir, timeout=timeout)
            elapsed = time.time() - t0

            result["agent_time"] = round(elapsed, 1)
            result["num_turns"] = data.get("num_turns")
            result["total_cost_usd"] = data.get("total_cost_usd")
            result["stop_reason"] = data.get("stop_reason")
            result["terminal_reason"] = data.get("terminal_reason")
            result["is_error"] = data.get("is_error", False)

            usage = data.get("usage", {})
            result["usage"] = {
                "input_tokens": usage.get("input_tokens", 0),
                "output_tokens": usage.get("output_tokens", 0),
                "cache_creation_input_tokens": usage.get("cache_creation_input_tokens", 0),
                "cache_read_input_tokens": usage.get("cache_read_input_tokens", 0),
            }

            model_usage = data.get("modelUsage", {})
            if model_usage:
                result["model_usage"] = model_usage

            agent_result = data.get("result", "")
            result["agent_result"] = agent_result[:1000] if agent_result else ""

            # Retry only on refusal-like failures.
            is_refusal = (
                result.get("stop_reason") == "refusal"
                or "violate our Usage Policy" in str(agent_result)
            )

            if is_refusal and attempt < max_retries:
                refusal_history.append({
                    "attempt": attempt,
                    "agent_time": result["agent_time"],
                    "num_turns": result["num_turns"],
                    "cost": result["total_cost_usd"],
                })
                last_result = result
                # Do not collect input or save result.json for an intermediate retry.
                continue

            # Success or final retry: collect input/.
            tmp_input = tmp_dir / "input"
            if tmp_input.exists():
                save_input = save_dir / "input"
                if save_input.exists():
                    shutil.rmtree(save_input)
                shutil.copytree(tmp_input, save_input)

            with open(save_dir / "claude_output.json", "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)

        except subprocess.TimeoutExpired:
            result["error"] = "timeout"
            result["agent_time"] = timeout
            if tmp_dir and (tmp_dir / "input").exists():
                save_input = save_dir / "input"
                if save_input.exists():
                    shutil.rmtree(save_input)
                shutil.copytree(tmp_dir / "input", save_input)
        except Exception as e:
            result["error"] = str(e)
        finally:
            if tmp_dir and tmp_dir.exists():
                shutil.rmtree(tmp_dir, ignore_errors=True)

        # Exit the retry loop.
        if refusal_history:
            result["refusal_retries"] = refusal_history
        with open(save_dir / "result.json", "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        return result

    # This should be unreachable because the final attempt always saves a result.
    if last_result is not None:
        last_result["refusal_retries"] = refusal_history
        with open(save_dir / "result.json", "w", encoding="utf-8") as f:
            json.dump(last_result, f, ensure_ascii=False, indent=2)
    return last_result


def collect_tasks(levels, qtypes, question_ids=None):
    """Collect task identifiers for the requested level/type/question filters."""
    tasks = []
    for level in levels:
        for qtype in qtypes:
            qtype_dir = DATA_DIR / f"level{level}" / qtype
            if not qtype_dir.exists():
                continue
            for q_dir in sorted(qtype_dir.iterdir()):
                if not q_dir.is_dir():
                    continue
                if not (q_dir / "task.json").exists():
                    continue
                qid = q_dir.name
                if question_ids and qid not in question_ids:
                    continue
                tasks.append((level, qtype, qid))
    return tasks


def main():
    global CLAUDE_CMD, CLAUDE_MODEL

    parser = argparse.ArgumentParser(description="OfficeEval Coding Agent (Claude Code CLI)")
    parser.add_argument("--level", type=int, choices=[1, 2], help="Task level to run")
    parser.add_argument("--type", choices=["word", "excel", "ppt"], help="Task type to run")
    parser.add_argument("--question", type=str, help="Question ID to run")
    parser.add_argument("--output-dir", type=str, help="Output directory")
    parser.add_argument("--timeout", type=int, default=3600, help="Timeout in seconds for each task (default: 3600)")
    parser.add_argument("--concurrency", type=int, default=1, help="Number of concurrent tasks (default: 1)")
    parser.add_argument("--claude-cmd", default=CLAUDE_CMD, help="Claude Code CLI command (default: claude)")
    parser.add_argument("--model", default=CLAUDE_MODEL, help="Claude model name")
    args = parser.parse_args()

    CLAUDE_CMD = args.claude_cmd
    CLAUDE_MODEL = args.model

    if shutil.which(CLAUDE_CMD) is None and not Path(CLAUDE_CMD).exists():
        raise SystemExit(
            f"Claude Code CLI not found: {CLAUDE_CMD}. "
            "Install Claude Code or pass --claude-cmd / set CLAUDE_CMD."
        )

    levels = [args.level] if args.level else [1, 2]
    qtypes = [args.type] if args.type else ["word", "excel", "ppt"]
    qids = [args.question] if args.question else None
    output_dir = args.output_dir or str(RESULTS_DIR / "coding-agent")

    # Clean up orphaned Claude CLI processes from previous runs before starting.
    if HAS_PSUTIL:
        killed = cleanup_orphan_claude_procs()
        if killed:
            print(f"[startup] Killed {killed} orphan claude process(es) from previous run")
    else:
        print("[startup] WARNING: psutil not installed, orphan cleanup + tree-kill disabled. Run: pip install psutil")

    tasks = collect_tasks(levels, qtypes, qids)

    # Resume support: skip tasks that already have result.json.
    remaining = []
    skipped = 0
    for level, qtype, qid in tasks:
        save_dir = Path(output_dir) / f"level{level}" / qtype / qid
        if (save_dir / "result.json").exists():
            skipped += 1
        else:
            remaining.append((level, qtype, qid))

    total = len(remaining)
    print(f"{'=' * 60}")
    print(f"OfficeEval Coding Agent (Claude Code)")
    print(f"{'=' * 60}")
    print(f"Model: {CLAUDE_MODEL}")
    print(f"Tasks: {total} ({skipped} skipped)")
    print(f"Output: {output_dir}")
    print(f"Timeout: {args.timeout}s per task")
    print(f"Concurrency: {args.concurrency}")
    print()

    if total == 0:
        print("All tasks already completed.")
        return

    results = []
    total_cost = 0.0
    completed = 0
    lock = threading.Lock()
    start_time = time.time()

    def run_and_report(task_info):
        level, qtype, qid = task_info
        return process_question(level, qtype, qid, output_dir, timeout=args.timeout)

    if args.concurrency <= 1:
        for i, (level, qtype, qid) in enumerate(remaining, 1):
            print(f"  [{i}/{total}] L{level} {qtype:5s} {qid} ...", end="", flush=True)

            result = run_and_report((level, qtype, qid))
            results.append(result)

            cost = result.get("total_cost_usd", 0) or 0
            total_cost += cost
            turns = result.get("num_turns", "?")
            t = result.get("agent_time", 0)
            error = result.get("error", "")
            status = f"ERR:{error[:50]}" if error else "OK"

            elapsed = time.time() - start_time
            eta = elapsed / i * (total - i)
            print(f"  {status}  ({turns} turns, ${cost:.2f}, {t:.0f}s, ETA {eta:.0f}s)")
    else:
        with ThreadPoolExecutor(max_workers=args.concurrency) as executor:
            future_to_task = {}
            for task_info in remaining:
                future = executor.submit(run_and_report, task_info)
                future_to_task[future] = task_info

            for future in as_completed(future_to_task):
                level, qtype, qid = future_to_task[future]
                try:
                    result = future.result()
                except Exception as e:
                    result = {"question_id": qid, "level": level, "type": qtype, "error": str(e)}

                with lock:
                    results.append(result)
                    completed += 1

                    cost = result.get("total_cost_usd", 0) or 0
                    total_cost += cost
                    turns = result.get("num_turns", "?")
                    t = result.get("agent_time", 0)
                    error = result.get("error", "")
                    status = f"ERR:{error[:50]}" if error else "OK"

                    elapsed = time.time() - start_time
                    eta = elapsed / completed * (total - completed)
                    print(f"  [{completed}/{total}] L{level} {qtype:5s} {qid}: {status}  ({turns} turns, ${cost:.2f}, {t:.0f}s, ETA {eta:.0f}s)")

    # Summary.
    total_time = time.time() - start_time
    errors = sum(1 for r in results if r.get("error"))
    avg_turns = sum(r.get("num_turns", 0) or 0 for r in results) / max(len(results), 1)

    print(f"\n{'=' * 60}")
    print(f"DONE in {total_time:.0f}s")
    print(f"  Completed: {len(results) - errors}/{len(results)}")
    print(f"  Errors: {errors}")
    print(f"  Avg turns: {avg_turns:.1f}")
    print(f"  Total cost: ${total_cost:.2f}")

    # Save run summary.
    summary_file = Path(output_dir) / f"agent_coding-agent_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    summary_file.parent.mkdir(parents=True, exist_ok=True)
    summary = {
        "agent": "coding-agent",
        "model": CLAUDE_MODEL,
        "timestamp": datetime.now().isoformat(),
        "total_tasks": total + skipped,
        "completed": len(results) - errors,
        "errors": errors,
        "total_cost_usd": round(total_cost, 2),
        "total_time_seconds": round(total_time, 1),
        "avg_turns": round(avg_turns, 1),
        "results": results,
    }
    with open(summary_file, "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    print(f"Summary: {summary_file}")


if __name__ == "__main__":
    main()
