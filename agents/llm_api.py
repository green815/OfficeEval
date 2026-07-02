#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Minimal LLM API adapter for the OfficeEval single-turn baseline.

Configuration is environment-variable driven. By default the adapter uses
OpenAI-compatible Chat Completions:

    OPENAI_API_KEY=...
    OPENAI_BASE_URL=https://api.openai.com/v1   # optional for OpenAI
    python agents/single_turn.py --model gpt-4.1

Useful overrides:

    OFFICEEVAL_API_KEY          API key, overrides OPENAI_API_KEY/API_KEY
    OFFICEEVAL_BASE_URL         base URL, overrides OPENAI_BASE_URL/API_BASE_URL
    OFFICEEVAL_PROVIDER         openai or anthropic
    OFFICEEVAL_API_FORMAT       chat or responses, for OpenAI-compatible APIs
    OFFICEEVAL_REASONING_EFFORT optional reasoning effort string
"""

import base64
import os
import time

import openai


def _env_first(*names):
    for name in names:
        value = os.environ.get(name)
        if value:
            return value
    return None


def _encode_image(img_path):
    with open(img_path, "rb") as f:
        img_bytes = f.read()
    img_data = base64.b64encode(img_bytes).decode("utf-8")
    if img_bytes[:4] == b"\x89PNG":
        media_type = "image/png"
    elif img_bytes[:2] == b"\xff\xd8":
        media_type = "image/jpeg"
    elif img_bytes[:4] == b"GIF8":
        media_type = "image/gif"
    elif img_bytes[:4] == b"RIFF" and img_bytes[8:12] == b"WEBP":
        media_type = "image/webp"
    else:
        media_type = "image/png"
    return {
        "type": "image_url",
        "image_url": {"url": f"data:{media_type};base64,{img_data}"},
    }


def _build_openai_content(user_content):
    content = []
    for item in user_content:
        if item["type"] == "text":
            content.append({"type": "text", "text": item["text"]})
        elif item["type"] == "image":
            content.append(_encode_image(item["path"]))
    return content


def _build_responses_content(openai_content):
    content = []
    for item in openai_content:
        if item["type"] == "text":
            content.append({"type": "input_text", "text": item["text"]})
        elif item["type"] == "image_url":
            content.append(
                {
                    "type": "input_image",
                    "image_url": item["image_url"]["url"],
                }
            )
    return content


def _usage_from_chat(completion):
    usage = getattr(completion, "usage", None)
    if not usage:
        return {}
    out = {
        "input_tokens": getattr(usage, "prompt_tokens", 0) or 0,
        "output_tokens": getattr(usage, "completion_tokens", 0) or 0,
    }
    details = getattr(usage, "completion_tokens_details", None)
    if details:
        reasoning = getattr(details, "reasoning_tokens", 0) or 0
        if reasoning:
            out["reasoning_tokens"] = reasoning
    return out


def _usage_from_responses(response):
    usage = getattr(response, "usage", None)
    if not usage:
        return {}
    out = {
        "input_tokens": getattr(usage, "input_tokens", 0) or 0,
        "output_tokens": getattr(usage, "output_tokens", 0) or 0,
    }
    details = getattr(usage, "output_tokens_details", None)
    if details:
        reasoning = getattr(details, "reasoning_tokens", 0) or 0
        if reasoning:
            out["reasoning_tokens"] = reasoning
    return out


def _call_openai(model_name, system_prompt, user_content):
    api_key = _env_first("OFFICEEVAL_API_KEY", "OPENAI_API_KEY", "API_KEY")
    base_url = _env_first("OFFICEEVAL_BASE_URL", "OPENAI_BASE_URL", "API_BASE_URL")
    api_format = os.environ.get("OFFICEEVAL_API_FORMAT", "chat").lower()
    reasoning_effort = os.environ.get("OFFICEEVAL_REASONING_EFFORT")
    timeout = float(os.environ.get("OFFICEEVAL_TIMEOUT", "1800"))

    client = openai.OpenAI(api_key=api_key, base_url=base_url, timeout=timeout)
    content = _build_openai_content(user_content)

    if api_format == "responses":
        request = {
            "model": model_name,
            "input": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": _build_responses_content(content)},
            ],
        }
        if reasoning_effort:
            request["reasoning"] = {"effort": reasoning_effort}
        response = client.responses.create(**request)
        return response.output_text or "", _usage_from_responses(response)

    request = {
        "model": model_name,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": content},
        ],
    }
    if reasoning_effort:
        request["reasoning_effort"] = reasoning_effort
    completion = client.chat.completions.create(**request)
    text = completion.choices[0].message.content or ""
    return text, _usage_from_chat(completion)


def _build_anthropic_content(user_content):
    content = []
    for item in _build_openai_content(user_content):
        if item["type"] == "text":
            content.append({"type": "text", "text": item["text"]})
        elif item["type"] == "image_url":
            url = item["image_url"]["url"]
            media_type = url.split(";")[0].split(":")[1]
            data = url.split(",", 1)[1]
            content.append(
                {
                    "type": "image",
                    "source": {
                        "type": "base64",
                        "media_type": media_type,
                        "data": data,
                    },
                }
            )
    return content


def _call_anthropic(model_name, system_prompt, user_content):
    try:
        import anthropic
    except ImportError as exc:
        raise ImportError("Install anthropic to use OFFICEEVAL_PROVIDER=anthropic") from exc

    api_key = _env_first("OFFICEEVAL_API_KEY", "ANTHROPIC_API_KEY")
    base_url = _env_first("OFFICEEVAL_BASE_URL", "ANTHROPIC_BASE_URL")
    max_tokens = int(os.environ.get("OFFICEEVAL_MAX_TOKENS", "32000"))
    timeout = float(os.environ.get("OFFICEEVAL_TIMEOUT", "1800"))

    client = anthropic.Anthropic(api_key=api_key, base_url=base_url, timeout=timeout)
    message = client.messages.create(
        model=model_name,
        max_tokens=max_tokens,
        system=system_prompt,
        messages=[{"role": "user", "content": _build_anthropic_content(user_content)}],
    )
    text_parts = []
    for block in message.content:
        if getattr(block, "type", "") == "text":
            text_parts.append(getattr(block, "text", "") or "")
    usage = getattr(message, "usage", None)
    usage_dict = {}
    if usage:
        usage_dict = {
            "input_tokens": getattr(usage, "input_tokens", 0) or 0,
            "output_tokens": getattr(usage, "output_tokens", 0) or 0,
        }
    return "".join(text_parts), usage_dict


def call_llm(model_name, system_prompt, user_content):
    """
    Return (response_text, usage_dict).

    user_content is a list containing:
      {"type": "text", "text": "..."}
      {"type": "image", "path": "..."}
    """
    provider = os.environ.get("OFFICEEVAL_PROVIDER")
    if not provider:
        provider = "anthropic" if model_name.startswith("claude-") else "openai"
    provider = provider.lower()

    if provider == "anthropic":
        fn = lambda: _call_anthropic(model_name, system_prompt, user_content)
    elif provider in {"openai", "custom"}:
        fn = lambda: _call_openai(model_name, system_prompt, user_content)
    else:
        raise ValueError(f"Unsupported OFFICEEVAL_PROVIDER: {provider}")

    max_retries = int(os.environ.get("OFFICEEVAL_MAX_RETRIES", "5"))
    for attempt in range(max_retries):
        try:
            return fn()
        except Exception:
            if attempt == max_retries - 1:
                raise
            wait = min(2 ** attempt * 5, 60)
            print(f"    [retry {attempt + 1}/{max_retries}] waiting {wait}s...")
            time.sleep(wait)
