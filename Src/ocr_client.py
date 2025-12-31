import base64
import datetime
import hashlib
import hmac
import json
import os
from typing import Optional
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlencode
from urllib.request import Request, urlopen

OCR_HOST_DEFAULT = "visual.volcengineapi.com"
OCR_REGION_DEFAULT = "cn-north-1"
OCR_SERVICE_DEFAULT = "cv"
OCR_IMAGE_MODE_DEFAULT = "base64"
OCR_MODE_DEFAULT = "default"
OCR_ACTION = "OCRNormal"
OCR_VERSION = "2020-08-26"


def get_env_value(name: str, default: Optional[str] = None) -> Optional[str]:
    value = os.environ.get(name)
    if value is None or value.strip() == "":
        return default
    return value.strip()


def resolve_visual_ocr_config() -> Optional[dict]:
    access_key = get_env_value("VOLC_ACCESS_KEY")
    secret_key = get_env_value("VOLC_SECRET_KEY")
    if not access_key or not secret_key:
        print("❌ 缺少 OCR 配置，请设置 VOLC_ACCESS_KEY、VOLC_SECRET_KEY")
        return None

    image_mode = get_env_value("OCR_IMAGE_MODE")
    if not image_mode:
        image_mode = get_env_value("IMAGEX_IMAGE_MODE", OCR_IMAGE_MODE_DEFAULT)

    image_url_prefix = get_env_value("OCR_IMAGE_URL_PREFIX")
    if not image_url_prefix:
        image_url_prefix = get_env_value("IMAGEX_IMAGE_URL_PREFIX")

    return {
        "access_key": access_key,
        "secret_key": secret_key,
        "region": get_env_value("OCR_REGION", OCR_REGION_DEFAULT),
        "service": get_env_value("OCR_SERVICE", OCR_SERVICE_DEFAULT),
        "host": get_env_value("OCR_ENDPOINT", OCR_HOST_DEFAULT),
        "image_mode": image_mode,
        "image_url_prefix": image_url_prefix,
        "approximate_pixel": get_env_value("OCR_APPROXIMATE_PIXEL"),
        "mode": get_env_value("OCR_MODE", OCR_MODE_DEFAULT),
        "filter_thresh": get_env_value("OCR_FILTER_THRESH"),
        "half_to_full": get_env_value("OCR_HALF_TO_FULL"),
        "session_token": get_env_value("VOLC_SESSION_TOKEN"),
    }


def resolve_ocr_workers() -> int:
    raw_value = get_env_value("OCR_MAX_WORKERS")
    if raw_value:
        try:
            value = int(raw_value)
            if value > 0:
                return value
        except ValueError:
            pass
    cpu_count = os.cpu_count() or 4
    return max(2, min(4, cpu_count))


def image_bytes_to_base64(image_bytes: bytes) -> str:
    return base64.b64encode(image_bytes).decode("ascii")


def image_path_to_base64(image_path) -> str:
    return image_bytes_to_base64(image_path.read_bytes())


def build_visual_ocr_body_from_base64(image_base64: str, config: dict) -> dict:
    body: dict = {"image_base64": image_base64}
    for key in ("approximate_pixel", "filter_thresh", "mode", "half_to_full"):
        value = config.get(key)
        if value:
            body[key] = value
    return body


def build_visual_ocr_body_from_path(image_path, config: dict) -> Optional[dict]:
    mode = (config["image_mode"] or OCR_IMAGE_MODE_DEFAULT).lower()
    if mode in {"image_url", "url"}:
        prefix = config["image_url_prefix"]
        if not prefix:
            print("❌ 缺少 OCR_IMAGE_URL_PREFIX，无法使用 image_url 模式")
            return None
        body = {"image_url": f"{prefix.rstrip('/')}/{image_path.name}"}
    else:
        body = {"image_base64": image_path_to_base64(image_path)}

    for key in ("approximate_pixel", "filter_thresh", "mode", "half_to_full"):
        value = config.get(key)
        if value:
            body[key] = value
    return body


def canonical_query(query: dict) -> str:
    items = []
    for key, value in query.items():
        items.append((quote(str(key), safe="-_.~"), quote(str(value), safe="-_.~")))
    return "&".join(f"{key}={value}" for key, value in sorted(items))


def hmac_sha256(key: bytes, msg: str) -> bytes:
    return hmac.new(key, msg.encode("utf-8"), hashlib.sha256).digest()


def get_signing_key(secret_key: str, date: str, region: str, service: str) -> bytes:
    kdate = hmac_sha256(secret_key.encode("utf-8"), date)
    kregion = hmac_sha256(kdate, region)
    kservice = hmac_sha256(kregion, service)
    return hmac_sha256(kservice, "request")


def sign_request(
    path: str,
    method: str,
    headers: dict,
    body: str,
    query: dict,
    access_key: str,
    secret_key: str,
    region: str,
    service: str,
    session_token: Optional[str] = None,
) -> None:
    if not path:
        path = "/"
    format_date = datetime.datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
    headers["X-Date"] = format_date

    body_hash = hashlib.sha256(body.encode("utf-8")).hexdigest()
    headers["X-Content-Sha256"] = body_hash
    if session_token:
        headers["X-Security-Token"] = session_token

    signed_headers = {}
    for key, value in headers.items():
        if key in {"Content-Type", "Content-Md5", "Host"} or key.startswith("X-"):
            signed_headers[key.lower()] = value

    if "host" in signed_headers:
        host_value = signed_headers["host"]
        if ":" in host_value:
            host, port = host_value.split(":", 1)
            if port in {"80", "443"}:
                signed_headers["host"] = host

    signed_header_lines = "".join(
        f"{key}:{signed_headers[key]}\n" for key in sorted(signed_headers)
    )
    signed_headers_string = ";".join(sorted(signed_headers))
    canonical_request = "\n".join(
        [
            method,
            path,
            canonical_query(query),
            signed_header_lines,
            signed_headers_string,
            body_hash,
        ]
    )
    credential_scope = "/".join([format_date[:8], region, service, "request"])
    signing_str = "\n".join(
        [
            "HMAC-SHA256",
            format_date,
            credential_scope,
            hashlib.sha256(canonical_request.encode("utf-8")).hexdigest(),
        ]
    )
    signing_key = get_signing_key(secret_key, format_date[:8], region, service)
    signature = hmac.new(signing_key, signing_str.encode("utf-8"), hashlib.sha256).hexdigest()
    headers["Authorization"] = (
        "HMAC-SHA256 "
        f"Credential={access_key}/{credential_scope}, "
        f"SignedHeaders={signed_headers_string}, "
        f"Signature={signature}"
    )


def request_visual_ocr(body_params: dict, config: dict) -> Optional[dict]:
    query = {"Action": OCR_ACTION, "Version": OCR_VERSION}
    body = urlencode(body_params)
    host = config["host"]
    headers = {
        "Host": host,
        "Content-Type": "application/x-www-form-urlencoded",
        "Accept": "application/json",
    }
    sign_request(
        "/",
        "POST",
        headers,
        body,
        query,
        config["access_key"],
        config["secret_key"],
        config["region"],
        config["service"],
        config.get("session_token"),
    )
    url = f"https://{host}/?{urlencode(query)}"
    request = Request(url, data=body.encode("utf-8"), headers=headers, method="POST")
    try:
        with urlopen(request, timeout=30) as response:
            data = response.read().decode("utf-8")
    except HTTPError as exc:
        data = exc.read().decode("utf-8") if exc.fp else ""
        print(f"❌ OCR 请求失败: HTTP {exc.code} {exc.reason}")
        if data:
            print(data)
        return None
    except URLError as exc:
        print(f"❌ OCR 请求失败: {exc}")
        return None
    try:
        return json.loads(data)
    except json.JSONDecodeError:
        print("❌ OCR 返回解析失败")
        return None


def extract_ocr_text(response: dict) -> str:
    if not response:
        return ""
    code = response.get("code")
    if code != 10000:
        message = response.get("message") or response.get("error") or ""
        if message:
            print(f"❌ OCR 返回错误: {message}")
        return ""
    data = response.get("data") or {}
    lines = [line for line in data.get("line_texts", []) if line]
    return "\n".join(lines).strip()


def ocr_image_path_to_text(image_path, config: Optional[dict] = None) -> str:
    config = config or resolve_visual_ocr_config()
    if not config:
        return ""
    body_params = build_visual_ocr_body_from_path(image_path, config)
    if not body_params:
        return ""
    response = request_visual_ocr(body_params, config)
    return extract_ocr_text(response) if response else ""


def ocr_image_bytes_to_text(image_bytes: bytes, config: Optional[dict] = None) -> str:
    config = config or resolve_visual_ocr_config()
    if not config:
        return ""
    body_params = build_visual_ocr_body_from_base64(image_bytes_to_base64(image_bytes), config)
    response = request_visual_ocr(body_params, config)
    return extract_ocr_text(response) if response else ""
