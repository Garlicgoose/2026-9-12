# -*- coding: utf-8 -*-
"""
FedEx Tracking API 模块

功能：
1. 使用 /track/v1/trackingnumbers 查询普通单或主单状态
2. 使用 /track/v1/associatedshipments 探测 MPS 关联单
3. 普通单以 trackingnumbers 接口结果为准
4. 多件货必须所有已返回关联单均为 Delivered，才判定整票送达
5. associatedshipments 返回达到 40 条时，增加人工复核备注
6. 整票送达后自动下载 Signature Proof of Delivery PDF
7. POD 文件名：主单号.pdf

POD 下载条件：
- save_pdf=True
- pdf_dir 不为空，或全局 PDF_DIR 已设置
- 运单/整票判断为 Delivered
- tracking API 返回 trackControlNumber 和 shipmentTimestamp

环境变量：
- FEDEX_API_KEY
- FEDEX_API_SECRET
- FEDEX_ACCOUNT_NUMBER（下载带签名 SPOD 时使用；未设置时默认 791310059）
"""

import base64
import binascii
import os
import sys
import time
import json
import random
import threading
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import requests


# ============================================================
# FedEx API 配置
# ============================================================

FEDEX_BASE_URL = "https://apis.fedex.com"
TOKEN_URL = f"{FEDEX_BASE_URL}/oauth/token"
TRACK_URL = f"{FEDEX_BASE_URL}/track/v1/trackingnumbers"
ASSOC_URL = f"{FEDEX_BASE_URL}/track/v1/associatedshipments"
DOC_URL = f"{FEDEX_BASE_URL}/track/v1/trackingdocuments"

LOCALE = "en_US"
REQUEST_TIMEOUT_SECONDS = 40
FEDEX_RELATED_LIMIT = 40
OVERWRITE_EXISTING_POD = True
FEDEX_ACCOUNT_NUMBER = os.getenv(
    "FEDEX_ACCOUNT_NUMBER",
    "791310059",
).strip()

# 稳定性配置
RETRYABLE_HTTP_STATUS = {429, 500, 502, 503, 504}
MAX_HTTP_ATTEMPTS = 5
BACKOFF_BASE_SECONDS = 1.0
BACKOFF_MAX_SECONDS = 30.0
BACKOFF_JITTER_SECONDS = 0.8
BATCH_REQUEST_INTERVAL_SECONDS = 0.5
FAILED_QUEUE_ROUNDS = 2
FAILED_QUEUE_DELAY_SECONDS = 5.0
STATUS_CACHE_FILE = Path(__file__).resolve().parent / "fedex_status_cache.json"

_shared_session: Optional[requests.Session] = None
_shared_session_lock = threading.Lock()
_cache_lock = threading.Lock()

DEFAULT_PDF_DIR = Path(__file__).resolve().parent / "pdf" / "FedEx"
PDF_DIR: Optional[Path] = DEFAULT_PDF_DIR

_token_cache: Dict[str, Any] = {
    "token": "",
    "expires": 0.0,
}


# ============================================================
# Session、重试与状态缓存
# ============================================================

def get_shared_session() -> requests.Session:
    """进程内复用同一个 Session，减少重复 TLS 连接。"""
    global _shared_session
    with _shared_session_lock:
        if _shared_session is None:
            _shared_session = requests.Session()
        return _shared_session


def close_shared_session() -> None:
    """程序退出前可主动调用；不调用也不影响正常运行。"""
    global _shared_session
    with _shared_session_lock:
        if _shared_session is not None:
            _shared_session.close()
            _shared_session = None


def _retry_delay(response: Optional[requests.Response], attempt: int) -> float:
    if response is not None:
        retry_after = response.headers.get("Retry-After", "").strip()
        if retry_after:
            try:
                return min(max(float(retry_after), 0.0), BACKOFF_MAX_SECONDS)
            except ValueError:
                pass
    exponential = BACKOFF_BASE_SECONDS * (2 ** max(attempt - 1, 0))
    return min(exponential + random.uniform(0.2, BACKOFF_JITTER_SECONDS), BACKOFF_MAX_SECONDS)


def _post_with_retry(
    session: requests.Session,
    url: str,
    *,
    max_attempts: int = MAX_HTTP_ATTEMPTS,
    **kwargs: Any,
) -> requests.Response:
    """对429、5xx、超时及连接错误执行指数退避重试。"""
    last_exception: Optional[BaseException] = None
    for attempt in range(1, max_attempts + 1):
        response: Optional[requests.Response] = None
        try:
            response = session.post(url, **kwargs)
            if response.status_code not in RETRYABLE_HTTP_STATUS:
                return response
            if attempt >= max_attempts:
                return response
        except (requests.Timeout, requests.ConnectionError) as exc:
            last_exception = exc
            if attempt >= max_attempts:
                raise
        time.sleep(_retry_delay(response, attempt))
    if last_exception is not None:
        raise last_exception
    raise RuntimeError("FedEx request failed without a response")


def _load_status_cache(cache_file: Any = None) -> Dict[str, Dict[str, Any]]:
    path = Path(cache_file or STATUS_CACHE_FILE)
    with _cache_lock:
        if not path.exists():
            return {}
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            return data if isinstance(data, dict) else {}
        except (OSError, ValueError, TypeError):
            return {}


def _save_status_cache(cache: Dict[str, Dict[str, Any]], cache_file: Any = None) -> None:
    path = Path(cache_file or STATUS_CACHE_FILE)
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = path.with_suffix(path.suffix + ".tmp")
    with _cache_lock:
        temp_path.write_text(
            json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        temp_path.replace(path)


def _is_temporary_failure(result: Dict[str, Any]) -> bool:
    text = " | ".join(
        str(result.get(key) or "") for key in ("error", "flag")
    ).casefold()
    markers = (
        "http 429", "http 500", "http 502", "http 503", "http 504",
        "service.unavailable", "too many requests", "network error",
        "timeout", "connectionerror", "invalid json",
    )
    return any(marker in text for marker in markers)


def _has_valid_tracking_status(result: Dict[str, Any]) -> bool:
    status = str(result.get("status") or "").strip().casefold()
    return bool(status) and status not in {
        "not found", "unknown", "temporary error", "rate limited", "network timeout"
    }


def _apply_cached_status(
    result: Dict[str, Any], cached: Optional[Dict[str, Any]]
) -> Dict[str, Any]:
    """临时失败时保留最后一次成功状态，同时保留本次错误信息。"""
    if not cached or not _is_temporary_failure(result):
        return result
    current_error = str(result.get("error") or "").strip()
    current_flag = str(result.get("flag") or "").strip()
    merged = dict(result)
    for key in (
        "status", "is_delivered", "delivery_date", "arrival_time",
        "pdf_file", "piece_count"
    ):
        if cached.get(key) not in (None, ""):
            merged[key] = cached[key]
    merged["from_cache"] = True
    merged["cache_updated_at"] = cached.get("cache_updated_at", "")
    merged["error"] = current_error
    merged["flag"] = current_flag
    return merged


# ============================================================
# 通用辅助函数
# ============================================================

def safe_print(message: Any) -> None:
    """安全输出，避免部分 Windows 终端编码错误。"""
    text = str(message)
    try:
        print(text)
    except UnicodeEncodeError:
        encoding = sys.stdout.encoding or "utf-8"
        print(text.encode(encoding, errors="replace").decode(encoding, errors="replace"))


def _api_headers(token: str) -> Dict[str, str]:
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
        "X-locale": LOCALE,
    }


def _extract_error_message(response: requests.Response) -> str:
    try:
        data = response.json()
    except ValueError:
        text = response.text.strip()
        return text[:1000] if text else "Unknown error"

    errors = data.get("errors") or []
    messages: List[str] = []

    if isinstance(errors, list):
        for error in errors:
            if not isinstance(error, dict):
                messages.append(str(error))
                continue
            code = str(error.get("code") or "").strip()
            message = str(error.get("message") or "").strip()
            if code and message:
                messages.append(f"{code}: {message}")
            elif code or message:
                messages.append(code or message)

    return "; ".join(messages) if messages else str(data)[:1000]


def _append_flag(result: Dict[str, Any], message: str) -> None:
    message = str(message or "").strip()
    if not message:
        return
    current = str(result.get("flag") or "").strip()
    result["flag"] = f"{current}；{message}" if current else message


def _get_token(
    api_key: str,
    api_secret: str,
    session: Optional[requests.Session] = None,
) -> str:
    now = time.time()
    if _token_cache["token"] and float(_token_cache["expires"]) > now + 60:
        return str(_token_cache["token"])

    http = session or requests.Session()
    response = _post_with_retry(http,
        TOKEN_URL,
        data={
            "grant_type": "client_credentials",
            "client_id": api_key,
            "client_secret": api_secret,
        },
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        timeout=REQUEST_TIMEOUT_SECONDS,
    )

    if response.status_code != 200:
        raise RuntimeError(
            f"FedEx auth failed (HTTP {response.status_code}): "
            f"{_extract_error_message(response)}"
        )

    try:
        data = response.json()
    except ValueError as exc:
        raise RuntimeError("FedEx auth returned invalid JSON") from exc

    token = str(data.get("access_token") or "").strip()
    if not token:
        raise RuntimeError("FedEx auth returned no access token")

    try:
        expires_in = float(data.get("expires_in", 3600))
    except (TypeError, ValueError):
        expires_in = 3600.0

    _token_cache["token"] = token
    _token_cache["expires"] = now + expires_in
    return token


# ============================================================
# 状态与时间处理
# ============================================================

def _fmt_delivery(value: str) -> str:
    if not value:
        return ""
    value = str(value).strip()
    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00")).strftime(
            "%Y-%m-%d %H:%M"
        )
    except ValueError:
        return value[:16] if len(value) >= 16 else value


def _latest_status(piece: Dict[str, Any]) -> Tuple[str, str]:
    if not isinstance(piece, dict):
        return "", ""
    latest = piece.get("latestStatusDetail") or {}
    code = str(latest.get("code") or "").strip()
    description = str(
        latest.get("statusByLocale") or latest.get("description") or ""
    ).strip()
    return code, description


def _is_delivered(piece: Dict[str, Any]) -> bool:
    code, description = _latest_status(piece)
    return code.upper() == "DL" or description.casefold() == "delivered"


def _actual_delivery_times(piece: Dict[str, Any]) -> List[str]:
    values: List[str] = []
    if not isinstance(piece, dict):
        return values

    for item in piece.get("dateAndTimes") or []:
        if not isinstance(item, dict):
            continue
        if str(item.get("type") or "").upper() == "ACTUAL_DELIVERY":
            value = str(item.get("dateTime") or "").strip()
            if value:
                values.append(value)

    if not values:
        latest = piece.get("latestStatusDetail") or {}
        value = str(latest.get("dateTime") or "").strip()
        if value:
            values.append(value)

    return values


def _latest_delivery_time(pieces: List[Dict[str, Any]]) -> str:
    values: List[str] = []
    for piece in pieces:
        values.extend(_actual_delivery_times(piece))

    if not values:
        return ""

    parsed: List[Tuple[float, str]] = []
    for value in values:
        try:
            dt = datetime.fromisoformat(value.replace("Z", "+00:00"))
            parsed.append((dt.timestamp(), value))
        except (ValueError, TypeError, OSError):
            continue

    if parsed:
        return max(parsed, key=lambda item: item[0])[1]
    return values[-1]


def _get_tracking_number(piece: Dict[str, Any]) -> str:
    if not isinstance(piece, dict):
        return ""
    return str(
        (piece.get("trackingNumberInfo") or {}).get("trackingNumber") or ""
    ).strip()


def _find_master(
    pieces: List[Dict[str, Any]], tracking_number: str
) -> Optional[Dict[str, Any]]:
    wanted = str(tracking_number).strip()
    for piece in pieces:
        if _get_tracking_number(piece) == wanted:
            return piece
    return pieces[0] if pieces else None


# ============================================================
# Tracking API
# ============================================================

def _query_trackingnumbers(
    session: requests.Session,
    token: str,
    tracking_number: str,
    timeout: int = REQUEST_TIMEOUT_SECONDS,
) -> Tuple[Optional[Dict[str, Any]], str]:
    payload = {
        "includeDetailedScans": False,
        "trackingInfo": [
            {"trackingNumberInfo": {"trackingNumber": str(tracking_number)}}
        ],
    }

    response = _post_with_retry(session,
        TRACK_URL,
        headers=_api_headers(token),
        json=payload,
        timeout=timeout,
    )

    if response.status_code != 200:
        return None, (
            f"tracking API HTTP {response.status_code}: "
            f"{_extract_error_message(response)}"
        )

    try:
        data = response.json()
    except ValueError:
        return None, "tracking API returned invalid JSON"

    complete_results = (data.get("output") or {}).get("completeTrackResults") or []
    if not complete_results:
        return None, "no completeTrackResults in tracking response"

    track_results = complete_results[0].get("trackResults") or []
    if not track_results:
        return None, "no trackResults in tracking response"

    wanted = str(tracking_number).strip()
    for piece in track_results:
        if _get_tracking_number(piece) == wanted:
            return piece, ""
    return track_results[0], ""


def _query_assoc(
    session: requests.Session,
    token: str,
    tracking_number: str,
    timeout: int = REQUEST_TIMEOUT_SECONDS,
) -> Tuple[List[Dict[str, Any]], str]:
    payload = {
        "includeDetailedScans": False,
        "associatedType": "STANDARD_MPS",
        "masterTrackingNumberInfo": {
            "trackingNumberInfo": {"trackingNumber": str(tracking_number)}
        },
    }

    response = _post_with_retry(session,
        ASSOC_URL,
        headers=_api_headers(token),
        json=payload,
        timeout=timeout,
    )

    if response.status_code != 200:
        return [], (
            f"assoc API HTTP {response.status_code}: "
            f"{_extract_error_message(response)}"
        )

    try:
        data = response.json()
    except ValueError:
        return [], "assoc API returned invalid JSON"

    complete_results = (data.get("output") or {}).get("completeTrackResults") or []
    if not complete_results:
        return [], "no completeTrackResults in assoc response"

    pieces = complete_results[0].get("trackResults") or []
    if not pieces:
        return [], "no trackResults in assoc response"
    return pieces, ""


# ============================================================
# POD 下载
# ============================================================

def _extract_document_content(document: Any) -> str:
    """兼容 documents 元素为 Base64 字符串或对象的情况。"""
    if isinstance(document, str):
        return document.strip()
    if not isinstance(document, dict):
        return ""

    for field in (
        "content",
        "document",
        "encodedContent",
        "encodedGraphic",
        "documentContent",
        "data",
    ):
        value = document.get(field)
        if value:
            return str(value).strip()
    return ""


def _decode_pdf_content(encoded_content: str) -> bytes:
    content = str(encoded_content or "").strip()
    if not content:
        raise ValueError("POD document content is empty")

    if content.startswith("data:") and "," in content:
        content = content.split(",", 1)[1]

    content = "".join(content.split())
    try:
        pdf_bytes = base64.b64decode(content, validate=True)
    except (ValueError, binascii.Error) as exc:
        raise ValueError("POD content is not valid Base64") from exc

    if not pdf_bytes:
        raise ValueError("Decoded POD PDF is empty")
    if not pdf_bytes.startswith(b"%PDF"):
        raise ValueError(
            "Decoded POD content is not a PDF. "
            f"First bytes: {pdf_bytes[:100]!r}"
        )
    return pdf_bytes


def _save_pod(
    session: requests.Session,
    tracking_number: str,
    token: str,
    master_piece: Dict[str, Any],
    pdf_dir: Any,
    timeout: int = REQUEST_TIMEOUT_SECONDS,
) -> Tuple[str, str]:
    """
    按已验证成功的 fedex_spod_download.py 流程下载签名 SPOD：
    1. 先用 trackingnumbers 查询当前请求运单；
    2. 从该运单响应读取 carrierCode 和 trackingNumberUniqueId；
    3. 在 trackingdocuments 请求中同时提交 billing accountNumber；
    4. 解码 output.documents[0] 并保存 PDF。

    返回：(pdf_file, error)
    """
    if not pdf_dir:
        return "", "POD directory is not set"

    tracking_number = str(tracking_number or "").strip()
    if not tracking_number:
        return "", "Tracking number is empty"
    if not FEDEX_ACCOUNT_NUMBER:
        return "", "FedEx billing account number is not set"

    output_directory = Path(pdf_dir).expanduser().resolve()
    output_directory.mkdir(parents=True, exist_ok=True)
    output_file = output_directory / f"{tracking_number}.pdf"

    # 关键：不能直接复用其他候选单号的 master_piece。
    # 必须像已测试成功的独立脚本一样，先查询当前请求号码，取得匹配的
    # carrierCode 与 trackingNumberUniqueId。
    exact_piece, tracking_error = _query_trackingnumbers(
        session=session,
        token=token,
        tracking_number=tracking_number,
        timeout=timeout,
    )
    source_piece = exact_piece or master_piece or {}
    piece_info = source_piece.get("trackingNumberInfo") or {}

    returned_number = str(piece_info.get("trackingNumber") or "").strip()
    carrier_code = str(piece_info.get("carrierCode") or "FDXE").strip()
    unique_id = str(piece_info.get("trackingNumberUniqueId") or "").strip()

    # trackingnumbers 若返回了另一条重复号码记录，不把不匹配的 Unique ID
    # 强行用于当前号码，以免获得无签名的普通 POD。
    if returned_number and returned_number != tracking_number:
        unique_id = ""

    tracking_number_info: Dict[str, str] = {
        "trackingNumber": tracking_number,
        "carrierCode": carrier_code or "FDXE",
    }
    if unique_id:
        tracking_number_info["trackingNumberUniqueId"] = unique_id

    payload = {
        "trackDocumentDetail": {
            "documentType": "SIGNATURE_PROOF_OF_DELIVERY",
            "documentFormat": "PDF",
        },
        "trackDocumentSpecification": [
            {
                "trackingNumberInfo": tracking_number_info,
                "accountNumber": FEDEX_ACCOUNT_NUMBER,
            }
        ],
    }

    document_headers = _api_headers(token)
    document_headers["x-customer-transaction-id"] = str(uuid.uuid4())

    try:
        response = _post_with_retry(
            session,
            DOC_URL,
            json=payload,
            headers=document_headers,
            timeout=timeout,
        )
    except requests.RequestException as exc:
        return "", f"POD request failed: {exc}"

    if response.status_code != 200:
        detail = _extract_error_message(response)
        if tracking_error:
            detail = f"{detail}; tracking lookup: {tracking_error}"
        return "", f"POD API HTTP {response.status_code}: {detail}"

    try:
        data = response.json()
    except ValueError:
        return "", f"POD API returned invalid JSON: {response.text[:500]}"

    output = data.get("output") or {}
    documents = output.get("documents") or []
    if not documents:
        return "", "POD API returned no documents. Response=" + str(data)[:1000]

    encoded_content = _extract_document_content(documents[0])
    if not encoded_content:
        return "", (
            "POD document has no PDF content. "
            f"type={type(documents[0]).__name__}"
        )

    try:
        pdf_bytes = _decode_pdf_content(encoded_content)
    except (ValueError, binascii.Error) as exc:
        return "", f"POD decode failed: {exc}"

    # 使用临时文件再替换，避免写入中断留下损坏 PDF。
    temp_file = output_file.with_suffix(".pdf.tmp")
    try:
        temp_file.write_bytes(pdf_bytes)
        temp_file.replace(output_file)
    except PermissionError as exc:
        # PDF 正在 Edge/Adobe/WPS 中打开时，Windows 可能禁止覆盖。
        # 不丢失本次成功下载结果，改存带时间戳的新文件。
        try:
            if temp_file.exists():
                temp_file.unlink()
        except OSError:
            pass
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        fallback_file = output_directory / f"{tracking_number}_{timestamp}.pdf"
        try:
            fallback_file.write_bytes(pdf_bytes)
        except OSError as fallback_exc:
            return "", (
                "POD save failed: target PDF is open or locked; "
                f"original={exc}; fallback={fallback_exc}"
            )
        output_file = fallback_file
    except OSError as exc:
        try:
            if temp_file.exists():
                temp_file.unlink()
        except OSError:
            pass
        return "", f"POD save failed: {exc}"

    if not output_file.exists() or output_file.stat().st_size == 0:
        return "", "POD PDF was not created or is empty"

    return str(output_file), ""

def _pod_request_numbers(
    tracking_number: str,
    preferred_piece: Optional[Dict[str, Any]],
    fallback_piece: Optional[Dict[str, Any]],
) -> List[str]:
    """决定 POD 下载尝试顺序：MPS 整票先试主单号，普通单用输入号。

    旧实现每次都拿「输入运单号」请求 POD（候选回退形同虚设）：
    当输入号是 MPS 子单时，用子单号请求签名 PDF 会失败；
    改为按 主单号 → 输入号 顺序逐个尝试，文件名 = 实际请求成功的运单号。
    """
    numbers: List[str] = []

    def _add(number: Any) -> None:
        number = str(number or "").strip()
        if number and number not in numbers:
            numbers.append(number)

    if isinstance(fallback_piece, dict):
        _add(_get_tracking_number(fallback_piece))
    if isinstance(preferred_piece, dict):
        _add(_get_tracking_number(preferred_piece))
    _add(tracking_number)
    return numbers


def _download_pod_to_result(
    result: Dict[str, Any],
    session: requests.Session,
    tracking_number: str,
    token: str,
    preferred_piece: Optional[Dict[str, Any]],
    fallback_piece: Optional[Dict[str, Any]],
    pdf_dir: Any,
) -> None:
    numbers = _pod_request_numbers(
        tracking_number, preferred_piece, fallback_piece
    )

    if not numbers:
        _append_flag(result, "POD未下载：没有可用的运单号")
        return

    errors: List[str] = []
    for number in numbers:
        pdf_file, error = _save_pod(
            session=session,
            tracking_number=number,
            token=token,
            master_piece=preferred_piece,
            pdf_dir=pdf_dir,
        )
        if pdf_file:
            result["pdf_file"] = pdf_file
            return
        if error:
            errors.append(error)

    unique_errors = list(dict.fromkeys(errors))
    _append_flag(result, "POD未下载：" + " | ".join(unique_errors))


# ============================================================
# 主查询函数
# ============================================================

def _query_fedex_one_live(
    tracking_number: str,
    api_key: str = "",
    api_secret: str = "",
    save_pdf: bool = True,
    pdf_dir: Any = None,
    _session: Optional[requests.Session] = None,
    _token: str = "",
) -> Dict[str, Any]:
    tracking_number = str(tracking_number or "").strip()
    result: Dict[str, Any] = {
        "tracking_number": tracking_number,
        "status": "",
        "is_delivered": "",
        "delivery_date": "",
        "arrival_time": "",
        "pdf_file": "",
        "error": "",
        "flag": "",
        "piece_count": 0,
        "from_cache": False,
        "cache_updated_at": "",
        "last_attempt_at": datetime.now(timezone.utc).isoformat(),
    }

    if not tracking_number:
        result["error"] = "Empty tracking number"
        return result
    if not api_key or not api_secret:
        result["error"] = "FedEx API key/secret not set"
        return result

    effective_pdf_dir = pdf_dir if pdf_dir is not None else PDF_DIR
    own_session = _session is None
    session = _session or requests.Session()

    try:
        token = _token or _get_token(api_key, api_secret, session=session)

        main_piece, main_error = _query_trackingnumbers(
            session, token, tracking_number
        )
        pieces, assoc_error = _query_assoc(session, token, tracking_number)
        if assoc_error or not pieces:
            pieces = []

        piece_count = len(pieces)
        result["piece_count"] = piece_count
        assoc_master = _find_master(pieces, tracking_number) if pieces else None
        master_piece = assoc_master or main_piece

        # 普通单或 associatedshipments 仅返回主单自身
        if piece_count <= 1:
            selected_piece = main_piece or master_piece
            if selected_piece is None:
                result["status"] = "Not Found"
                result["error"] = main_error or assoc_error or "No track results"
                return result

            code, description = _latest_status(selected_piece)
            result["status"] = description or code or "Unknown"

            if _is_delivered(selected_piece):
                result["status"] = "Delivered"
                result["is_delivered"] = "Y"
                delivery_time = _latest_delivery_time([selected_piece])
                if delivery_time:
                    formatted = _fmt_delivery(delivery_time)
                    result["delivery_date"] = formatted
                    result["arrival_time"] = formatted

                if save_pdf:
                    _download_pod_to_result(
                        result=result,
                        session=session,
                        tracking_number=tracking_number,
                        token=token,
                        preferred_piece=main_piece,
                        fallback_piece=assoc_master,
                        pdf_dir=effective_pdf_dir,
                    )
            return result

        # 多件货：检查所有 associatedshipments 返回件
        undelivered: List[Tuple[str, str]] = []
        for piece in pieces:
            if not _is_delivered(piece):
                code, description = _latest_status(piece)
                undelivered.append(
                    (
                        _get_tracking_number(piece) or "Unknown",
                        description or code or "Unknown",
                    )
                )

        if piece_count >= FEDEX_RELATED_LIMIT:
            _append_flag(
                result,
                "关联单返回达到40条接口上限，可能存在未返回的隐藏子单，需人工复核",
            )

        if undelivered:
            result["status"] = undelivered[0][1]
            shown = undelivered[:5]
            detail = "; ".join(f"{number}={status}" for number, status in shown)
            if len(undelivered) > 5:
                detail += "..."
            _append_flag(result, f"未送达关联单{len(undelivered)}件：{detail}")
            return result

        result["status"] = "Delivered"
        result["is_delivered"] = "Y"
        delivery_time = _latest_delivery_time(pieces)
        if delivery_time:
            formatted = _fmt_delivery(delivery_time)
            result["delivery_date"] = formatted
            result["arrival_time"] = formatted

        if piece_count >= FEDEX_RELATED_LIMIT:
            _append_flag(result, "当前返回的40个关联单均已送达")

        if save_pdf:
            _download_pod_to_result(
                result=result,
                session=session,
                tracking_number=tracking_number,
                token=token,
                preferred_piece=main_piece,
                fallback_piece=master_piece,
                pdf_dir=effective_pdf_dir,
            )
        return result

    except requests.RequestException as exc:
        result["error"] = f"Network error: {exc}"
        return result
    except Exception as exc:
        result["error"] = f"{type(exc).__name__}: {exc}"
        return result
    finally:
        if own_session:
            session.close()


# ============================================================
# 带缓存的公开查询函数与失败队列批量查询
# ============================================================

def query_fedex_one(
    tracking_number: str,
    api_key: str = "",
    api_secret: str = "",
    save_pdf: bool = True,
    pdf_dir: Any = None,
    session: Optional[requests.Session] = None,
    token: str = "",
    use_cache: bool = True,
    cache_file: Any = None,
) -> Dict[str, Any]:
    """查询单票；临时失败时自动回退到最后一次成功状态。"""
    tracking_number = str(tracking_number or "").strip()
    cache = _load_status_cache(cache_file) if use_cache else {}
    # 单票公开入口也复用进程级 Session。主运行器逐行调用时不再重复
    # 建立 TLS 连接，批量查询会明显更稳定。
    effective_session = session or get_shared_session()
    result = _query_fedex_one_live(
        tracking_number=tracking_number,
        api_key=api_key,
        api_secret=api_secret,
        save_pdf=save_pdf,
        pdf_dir=pdf_dir,
        _session=effective_session,
        _token=token,
    )

    temporary_failure = _is_temporary_failure(result)
    if temporary_failure:
        return _apply_cached_status(result, cache.get(tracking_number))

    if use_cache and _has_valid_tracking_status(result):
        result["cache_updated_at"] = datetime.now(timezone.utc).isoformat()
        cache[tracking_number] = {
            key: result.get(key)
            for key in (
                "tracking_number", "status", "is_delivered", "delivery_date",
                "arrival_time", "pdf_file", "piece_count", "cache_updated_at"
            )
        }
        try:
            _save_status_cache(cache, cache_file)
        except OSError:
            pass  # 缓存目录不可写（打包环境）时静默跳过，不影响查询
    return result


def query_fedex_batch(
    tracking_numbers: List[str],
    api_key: str = "",
    api_secret: str = "",
    save_pdf: bool = True,
    pdf_dir: Any = None,
    cache_file: Any = None,
    failed_queue_rounds: int = FAILED_QUEUE_ROUNDS,
    request_interval: float = BATCH_REQUEST_INTERVAL_SECONDS,
) -> List[Dict[str, Any]]:
    """
    批量查询并复用Session/Token。
    第一轮结束后，仅对临时失败项执行失败队列重试。
    """
    numbers = list(dict.fromkeys(
        str(number or "").strip() for number in tracking_numbers if str(number or "").strip()
    ))
    if not numbers:
        return []

    session = get_shared_session()
    token = _get_token(api_key, api_secret, session=session)
    results: Dict[str, Dict[str, Any]] = {}
    retry_queue = list(numbers)

    for round_index in range(max(0, failed_queue_rounds) + 1):
        current_queue = retry_queue
        retry_queue = []
        if round_index > 0:
            time.sleep(FAILED_QUEUE_DELAY_SECONDS * round_index)

        for index, tracking_number in enumerate(current_queue):
            result = query_fedex_one(
                tracking_number=tracking_number,
                api_key=api_key,
                api_secret=api_secret,
                save_pdf=save_pdf,
                pdf_dir=pdf_dir,
                session=session,
                token=token,
                use_cache=True,
                cache_file=cache_file,
            )
            result["queue_round"] = round_index + 1
            results[tracking_number] = result

            # 即使回退到缓存，只要本次API仍临时失败，就继续进入失败队列。
            if _is_temporary_failure(result):
                retry_queue.append(tracking_number)
            if request_interval > 0 and index < len(current_queue) - 1:
                time.sleep(request_interval)

        if not retry_queue:
            break

    return [results[number] for number in numbers]


# ============================================================
# 命令行测试
# ============================================================

def main() -> int:
    """
    PowerShell：
        $env:FEDEX_API_KEY="你的API Key"
        $env:FEDEX_API_SECRET="你的API Secret"
        python fedex_module_with_pod.py 472887355535

    指定保存目录：
        python fedex_module_with_pod.py 472887355535 "D:\\FedEx_POD"
    """
    tracking_number = sys.argv[1] if len(sys.argv) > 1 else ""
    if not tracking_number:
        safe_print("Usage: python fedex_module_with_pod.py <tracking_number> [pdf_dir]")
        return 1

    selected_pdf_dir = Path(sys.argv[2]) if len(sys.argv) > 2 else DEFAULT_PDF_DIR
    api_key = os.getenv("FEDEX_API_KEY", "").strip()
    api_secret = os.getenv("FEDEX_API_SECRET", "").strip()

    if not api_key or not api_secret:
        safe_print("Error: FEDEX_API_KEY or FEDEX_API_SECRET is not set.")
        return 1

    result = query_fedex_one(
        tracking_number=tracking_number,
        api_key=api_key,
        api_secret=api_secret,
        save_pdf=True,
        pdf_dir=selected_pdf_dir,
    )

    safe_print("=" * 70)
    for key, value in result.items():
        safe_print(f"{key}: {value}")
    safe_print("=" * 70)

    if result["pdf_file"]:
        safe_print("POD downloaded successfully:")
        safe_print(result["pdf_file"])
    elif result["is_delivered"] == "Y":
        safe_print("Shipment is delivered, but POD was not downloaded.")
        safe_print("Please check the flag field for the reason.")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
