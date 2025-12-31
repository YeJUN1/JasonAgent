import os
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional

from volcenginesdkarkruntime import Ark

DEFAULT_BASE_URL = "https://ark.cn-beijing.volces.com/api/v3"
DEFAULT_MODEL = "doubao-seed-1-6-251015"


def build_messages(text: str, image_url: Optional[str] = None) -> List[Dict[str, Any]]:
    if image_url:
        return [
            {
                "role": "user",
                "content": [
                    {"type": "image_url", "image_url": {"url": image_url}},
                    {"type": "text", "text": text},
                ],
            }
        ]
    return [{"role": "user", "content": text}]


def create_client(
    api_key: Optional[str] = None,
    base_url: str = DEFAULT_BASE_URL,
) -> Ark:
    if not os.environ.get("ARK_API_KEY"):
        load_env_file()
    key = api_key or os.environ.get("ARK_API_KEY")
    if not key:
        raise RuntimeError("Missing ARK_API_KEY environment variable.")
    return Ark(base_url=base_url, api_key=key)


def load_env_file(env_path: Optional[Path] = None) -> None:
    path = env_path or (Path(__file__).resolve().parent.parent / ".env")
    if not path.exists():
        return
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


def chat_completion(
    messages: List[Dict[str, Any]],
    model: str = DEFAULT_MODEL,
    base_url: str = DEFAULT_BASE_URL,
    api_key: Optional[str] = None,
    reasoning_effort: str = "medium",
) -> str:
    client = create_client(api_key=api_key, base_url=base_url)
    completion = client.chat.completions.create(
        model=model,
        messages=messages,
        reasoning_effort=reasoning_effort,
    )
    return completion.choices[0].message.content or ""


def chat_completion_stream(
    messages: List[Dict[str, Any]],
    model: str = DEFAULT_MODEL,
    base_url: str = DEFAULT_BASE_URL,
    api_key: Optional[str] = None,
    reasoning_effort: str = "medium",
) -> Iterable[str]:
    client = create_client(api_key=api_key, base_url=base_url)
    stream = client.chat.completions.create(
        model=model,
        messages=messages,
        reasoning_effort=reasoning_effort,
        stream=True,
    )
    for chunk in stream:
        if not chunk.choices:
            continue
        delta = chunk.choices[0].delta.content
        if delta:
            yield delta


if __name__ == "__main__":
    sample_messages = build_messages("Hello from Doubao.")
    print(chat_completion(sample_messages))
