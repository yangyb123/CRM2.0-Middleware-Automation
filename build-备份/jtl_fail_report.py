#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CRM2.0 中间件接口 JTL 失败报告生成器

功能:
  - 读取 JMeter 生成的 .jtl 文件，默认支持 CSV JTL，也兼容常见 XML JTL。
  - 只输出失败/报错的采样结果。
  - 从 failureMessage 中提取接口、字段、预期值(DB/CSV)、实际值(API)、DB字段、辅助接口值、失败原因。
  - 生成一个可直接打开的 HTML 报告，样式内嵌，无外部依赖。

使用:
  python jtl_fail_report.py
  python jtl_fail_report.py 中间件接口结果.jtl
  python jtl_fail_report.py 中间件接口结果.jtl -o 中间件接口对比失败报告.html
  python jtl_fail_report.py 中间件接口结果.jtl --config jtl_fail_report_config.json --no-open
  python jtl_fail_report.py 中间件接口结果.jtl --no-open --teams-send-file --teams-graph-token <GraphToken>
  python jtl_fail_report.py 中间件接口结果.jtl --no-open --teams-send-file --teams-auto-login
"""

import argparse
import csv
import glob
import io
import json
import os
import re
import sys
import time
import uuid
import xml.etree.ElementTree as ET
from datetime import datetime
from html import escape
from pathlib import Path
from urllib import error as urlerror
from urllib.parse import parse_qs, parse_qsl, quote, unquote, urlencode, urlparse
from urllib import request as urlrequest


SCRIPT_VERSION = "4.19"
DEFAULT_CONFIG_FILE = "jtl_fail_report_config.json"
DEFAULT_TOKEN_CACHE_FILE = "jtl_fail_report_token_cache.json"
DEFAULT_TEAMS_CHAT_URL = "https://teams.microsoft.com/l/chat/19:4f1f0b60d3a94f90b0c883360575d0d8@thread.v2/conversations?context=%7B%22contextType%22%3A%22chat%22%7D"
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
GRAPH_DEFAULT_SCOPES = "offline_access https://graph.microsoft.com/Files.ReadWrite https://graph.microsoft.com/ChatMessage.Send"


def h(value):
    """HTML escape helper."""
    if value is None:
        return ""
    return escape(str(value), quote=True)


def read_text(path):
    """Read text with a practical fallback for files saved by Excel/JMeter on Windows."""
    for enc in ("utf-8-sig", "utf-8", "gbk"):
        try:
            with open(path, "r", encoding=enc, newline="") as f:
                return f.read()
        except UnicodeDecodeError:
            continue
    with open(path, "r", encoding="utf-8", errors="replace", newline="") as f:
        return f.read()


XML_CHAR_REF_RE = re.compile(r"&#(x[0-9A-Fa-f]+|X[0-9A-Fa-f]+|[0-9]+);")


def is_valid_xml_codepoint(codepoint):
    return (
        codepoint in (0x9, 0xA, 0xD)
        or 0x20 <= codepoint <= 0xD7FF
        or 0xE000 <= codepoint <= 0xFFFD
        or 0x10000 <= codepoint <= 0x10FFFF
    )


def sanitize_xml_text(text):
    """Remove XML 1.0-invalid characters/references so one bad response cannot block the report."""
    stats = {"invalid_char_refs": 0, "invalid_raw_chars": 0}

    def replace_char_ref(match):
        raw = match.group(1)
        try:
            if raw[:1].lower() == "x":
                codepoint = int(raw[1:], 16)
            else:
                codepoint = int(raw, 10)
        except ValueError:
            stats["invalid_char_refs"] += 1
            return "\ufffd"
        if is_valid_xml_codepoint(codepoint):
            return match.group(0)
        stats["invalid_char_refs"] += 1
        return "\ufffd"

    text = XML_CHAR_REF_RE.sub(replace_char_ref, text)
    cleaned = []
    for ch in text:
        if is_valid_xml_codepoint(ord(ch)):
            cleaned.append(ch)
        else:
            stats["invalid_raw_chars"] += 1
            cleaned.append("\ufffd")
    return "".join(cleaned), stats


def parse_xml_root(text):
    try:
        return ET.fromstring(text)
    except ET.ParseError as exc:
        cleaned, stats = sanitize_xml_text(text)
        if cleaned == text:
            raise
        print(
            "提示: XML JTL 中存在非法 XML 字符，已清理后继续生成报告；"
            f"非法字符引用 {stats['invalid_char_refs']} 个，"
            f"非法原始字符 {stats['invalid_raw_chars']} 个。原始错误: {exc}"
        )
        return ET.fromstring(cleaned)


def find_jtl_file(specified_path=None):
    """Use explicit input first; otherwise pick the newest .jtl in the script directory."""
    if specified_path:
        if os.path.isfile(specified_path):
            return os.path.abspath(specified_path)
        print(f"指定的文件不存在: {specified_path}")
        sys.exit(1)

    script_dir = os.path.dirname(os.path.abspath(__file__))
    jtl_files = glob.glob(os.path.join(script_dir, "*.jtl"))
    if not jtl_files:
        print(f"当前目录下未找到 .jtl 文件: {script_dir}")
        print("请将脚本放到 .jtl 文件所在目录，或指定 JTL 文件路径运行。")
        sys.exit(1)

    jtl_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    if len(jtl_files) > 1:
        print(f"发现 {len(jtl_files)} 个 JTL 文件，默认选择最新的:")
        for i, item in enumerate(jtl_files[:5], 1):
            mtime = datetime.fromtimestamp(os.path.getmtime(item)).strftime("%Y-%m-%d %H:%M:%S")
            size = os.path.getsize(item) / 1024
            mark = " <- selected" if i == 1 else ""
            print(f"  {i}. {os.path.basename(item)} ({size:.1f} KB, {mtime}){mark}")
    return jtl_files[0]


def parse_jtl(path):
    """Parse CSV JTL or common XML JTL into a list of row dictionaries."""
    text = read_text(path)
    if text.lstrip().startswith("<"):
        return parse_xml_jtl(text)
    reader = csv.DictReader(io.StringIO(text))
    return list(reader)


def child_text_by_local_name(elem, local_names):
    wanted = set(local_names)
    for child in list(elem):
        tag = child.tag.split("}")[-1]
        if tag in wanted:
            return child.text or ""
    return ""


def parse_xml_jtl(text):
    rows = []
    root = parse_xml_root(text)
    for elem in root.iter():
        tag = elem.tag.split("}")[-1]
        if tag not in ("httpSample", "sample"):
            continue
        attrs = elem.attrib
        method = child_text_by_local_name(elem, ("method", "Method", "requestMethod"))
        query_string = child_text_by_local_name(elem, ("queryString", "QueryString", "query_string"))
        sampler_data = child_text_by_local_name(elem, ("samplerData", "SamplerData", "requestData", "RequestData"))
        assertion_messages = []
        for ar in elem.findall(".//assertionResult"):
            failure = (ar.findtext("failure") or "").strip().lower()
            error = (ar.findtext("error") or "").strip().lower()
            msg = ar.findtext("failureMessage") or ""
            if failure == "true" or error == "true" or msg.strip():
                assertion_messages.append(msg)
        url = attrs.get("url", "") or child_text_by_local_name(elem, ("java.net.URL", "URL", "url"))
        rows.append({
            "timeStamp": attrs.get("ts", ""),
            "elapsed": attrs.get("t", ""),
            "label": attrs.get("lb", ""),
            "responseCode": attrs.get("rc", ""),
            "responseMessage": attrs.get("rm", ""),
            "threadName": attrs.get("tn", ""),
            "success": "true" if attrs.get("s", "true").lower() == "true" else "false",
            "failureMessage": "\n".join(m for m in assertion_messages if m),
            "responseData": child_text_by_local_name(elem, ("responseData", "responseBody", "response_body")),
            "method": method,
            "queryString": query_string,
            "requestData": query_string or sampler_data,
            "URL": url,
        })
    return rows


def first_match(patterns, text, default=""):
    for pat in patterns:
        m = re.search(pat, text, re.S)
        if m:
            value = (m.group(1) or "").strip().strip(",")
            if value:
                return value
    return default


def query_value(url, key):
    if not url or url == "null":
        return ""
    try:
        parsed = urlparse(url)
        values = parse_qs(parsed.query).get(key)
        return values[0] if values else ""
    except Exception:
        return ""


def endpoint_from_crm20_text(text):
    if not text:
        return ""
    m = re.search(r"(?:^|[：:\s])/?(crm20/[A-Za-z0-9_]+)(?:[/?#\s]|$)", text)
    return m.group(1) if m else ""


def extract_api_name(label, url):
    if url and url != "null":
        m = re.search(r"/crm20/([A-Za-z0-9_]+)", url)
        if m:
            return m.group(1)
    m = re.search(r"\b(get_[A-Za-z0-9_]+|modify_user_info|change_user_rights|set_group_permission)\b", label or "")
    if m:
        return m.group(1)
    label_endpoint = endpoint_from_crm20_text(label)
    if label_endpoint:
        return label_endpoint
    mapping = [
        ("挂单", "get_order_info"),
        ("货币", "get_symbols_info"),
        ("符号规格", "get_symbol_ex"),
        ("实时报价", "get_symbol_ex"),
        ("持仓", "get_position_ex"),
        ("已平仓", "get_history_position_ex"),
        ("汇总", "get_history_position_ex"),
        ("财务指标", "get_account_metrics_ex"),
        ("历史订单", "get_order_ex"),
        ("待处理", "get_order_ex"),
        ("用户索引", "get_all_users"),
        ("设置组权限", "set_group_permission"),
    ]
    for keyword, api in mapping:
        if keyword in label:
            return api
    return label or "unknown"


def response_data_from_row(row):
    if not row:
        return ""
    preferred = (
        "responseData",
        "responseDataAsString",
        "responseBody",
        "response_body",
        "response_data",
        "ResponseData",
        "ResponseBody",
        "response",
    )
    for key in preferred:
        value = row.get(key)
        if value:
            return str(value)
    lower_map = {str(k).strip().lower(): v for k, v in row.items()}
    for key in preferred:
        value = lower_map.get(key.lower())
        if value:
            return str(value)
    return ""


def request_method_from_row(row):
    if not row:
        return ""
    preferred = (
        "method",
        "requestMethod",
        "httpMethod",
        "RequestMethod",
        "HTTPMethod",
    )
    for key in preferred:
        value = row.get(key)
        if value:
            return str(value).strip().upper()
    lower_map = {str(k).strip().lower(): v for k, v in row.items()}
    for key in preferred:
        value = lower_map.get(key.lower())
        if value:
            return str(value).strip().upper()
    return ""


def request_data_from_row(row):
    if not row:
        return ""
    preferred = (
        "requestData",
        "queryString",
        "requestBody",
        "request_body",
        "RequestData",
        "RequestBody",
        "samplerData",
        "SamplerData",
    )
    for key in preferred:
        value = row.get(key)
        if value:
            return str(value)
    lower_map = {str(k).strip().lower(): v for k, v in row.items()}
    for key in preferred:
        value = lower_map.get(key.lower())
        if value:
            return str(value)
    return ""


def looks_like_json_payload(text):
    text = (text or "").strip()
    if text.startswith("{"):
        return True
    if not text.startswith("["):
        return False
    return bool(re.match(r"^\[\s*(?:[\{\[\"0-9tfn\-]|]|$)", text, re.I))


def looks_like_compare_report(text):
    text = (text or "").strip()
    if not text:
        return False
    markers = (
        "Assertion message:",
        "字段对比结果",
        "对比结果",
        "成功明细",
        "失败明细",
        "跳过字段明细",
    )
    return any(marker in text for marker in markers)


def looks_like_non_http_sampler_data(text):
    text = (text or "").strip()
    if not text:
        return False
    if looks_like_json_payload(text):
        return False
    prefixes = (
        "import ",
        "String ",
        "File ",
        "def ",
        "[Prepared Select Statement]",
        "[Select Statement]",
        "SELECT ",
        "WITH ",
    )
    if text.startswith(prefixes):
        return True
    return "org.apache.jmeter" in text or "JsonSlurper" in text or "vars.get" in text


def compact_display_value(value, limit=500):
    if value is None:
        text = "null"
    else:
        try:
            text = json.dumps(value, ensure_ascii=False, separators=(",", ":"))
        except Exception:
            text = str(value)
    if limit and len(text) > limit:
        return text[: max(limit - 3, 0)] + "..."
    return text


def api_code_is_success(code):
    if code is None:
        return True
    if isinstance(code, bool):
        return code
    if isinstance(code, (int, float)):
        return code == 0
    text = str(code).strip().strip('"').strip("'").lower()
    if not text:
        return True
    return text in ("0", "0.0", "ok", "success", "true")


def api_value_is_empty(value):
    if value is None:
        return True
    if isinstance(value, str):
        text = value.strip().lower()
        return text in ("", "null", "[]", "{}")
    if isinstance(value, (list, tuple, set)):
        return len(value) == 0
    if isinstance(value, dict):
        return len(value) == 0
    return False


def find_empty_api_data_node(parsed):
    if isinstance(parsed, dict):
        if "data" in parsed:
            data = parsed.get("data")
            if api_value_is_empty(data):
                return "data", data
            if isinstance(data, dict):
                for key in ("items", "rows", "list", "records", "orders", "positions", "deals"):
                    if key in data and api_value_is_empty(data.get(key)):
                        return f"data.{key}", data.get(key)
        return "", None
    if isinstance(parsed, list) and not parsed:
        return "response", parsed
    return "", None


def message_indicates_api_empty_data(message):
    text = re.sub(r"\s+", "", message or "").lower()
    if not text:
        return False
    markers = (
        "接口没有返回",
        "接口未返回",
        "未返回有效订单号",
        "没有返回可对比",
        "无可对比",
        "data为空",
        "data返回为空",
        "data无数据",
        "date无数据",
        "响应data为空",
        "响应date为空",
    )
    return any(marker.lower() in text for marker in markers)


def detect_api_response_issues(info):
    response_data = info.get("response_data") or ""
    message = info.get("failure_message") or ""
    parsed = parse_json_response(response_data)
    issues = []

    if isinstance(parsed, dict):
        code = parsed.get("code")
        msg = parsed.get("msg", parsed.get("message", parsed.get("error", "")))
        meta = parsed.get("meta", "")
        if code is not None and not api_code_is_success(code):
            issues.append({
                "type": "API_ERROR",
                "title": "接口业务失败",
                "node": "code/msg",
                "actual_value": f"code={compact_display_value(code, 120)}, msg={compact_display_value(msg, 260)}",
                "expected_value": "code=0 且 msg=ok/success",
                "code": compact_display_value(code, 120),
                "msg": compact_display_value(msg, 260),
                "meta": compact_display_value(meta, 260),
                "reason": "接口返回业务失败码，不能继续做字段对比。",
            })

        success_value = parsed.get("success")
        if success_value is False:
            msg = parsed.get("msg", parsed.get("message", parsed.get("error", "")))
            issues.append({
                "type": "API_ERROR",
                "title": "接口业务失败",
                "node": "success",
                "actual_value": "false",
                "expected_value": "true",
                "code": compact_display_value(code, 120),
                "msg": compact_display_value(msg, 260),
                "meta": compact_display_value(parsed.get("meta", ""), 260),
                "reason": "接口响应 success=false，不能继续做字段对比。",
            })

        empty_node, empty_value = find_empty_api_data_node(parsed)
        if empty_node:
            issues.append({
                "type": "API_EMPTY_DATA",
                "title": "接口响应 data 为空",
                "node": empty_node,
                "actual_value": compact_display_value(empty_value, 260),
                "expected_value": "非空数组/对象，并包含可比对记录",
                "code": compact_display_value(code, 120),
                "msg": compact_display_value(msg, 260),
                "meta": compact_display_value(parsed.get("meta", ""), 260),
                "reason": "接口响应没有返回可用于对比的业务数据；请检查 instanceid、login、ticket/order、时间范围或数据库是否有对应记录。",
            })
    elif isinstance(parsed, list) and not parsed:
        issues.append({
            "type": "API_EMPTY_DATA",
            "title": "接口响应为空数组",
            "node": "response",
            "actual_value": "[]",
            "expected_value": "非空数组/对象，并包含可比对记录",
            "code": "",
            "msg": "",
            "meta": "",
            "reason": "接口响应为空数组，没有可用于字段对比的记录。",
        })

    if message_indicates_api_empty_data(message) and not any(x.get("type") == "API_EMPTY_DATA" for x in issues):
        issues.append({
            "type": "API_EMPTY_DATA",
            "title": "接口未返回有效业务数据",
            "node": "data",
            "actual_value": compact_display_value(parsed, 260) if parsed is not None else "<未解析到JSON响应，见接口响应正文>",
            "expected_value": "接口返回有效订单号/仓位号及可比对记录",
            "code": "",
            "msg": "",
            "meta": "",
            "reason": first_non_empty_segment(message) or "接口未返回有效业务数据，无法继续字段对比。",
        })

    deduped = []
    seen = set()
    for issue in issues:
        key = (issue.get("type"), issue.get("node"), issue.get("actual_value"), issue.get("reason"))
        if key in seen:
            continue
        seen.add(key)
        deduped.append(issue)
    return deduped


def api_context_from_row(row):
    """Capture the real HTTP sampler payload so a later JSR223 failure can reuse it."""
    url = (row.get("URL") or "").strip()
    response_data = response_data_from_row(row)
    if not url or url == "null" or "/crm20/" not in url:
        return None
    if not looks_like_json_payload(response_data):
        return None
    return {
        "url": url,
        "method": request_method_from_row(row),
        "request_data": request_data_from_row(row),
        "response_data": response_data,
        "api_name": extract_api_name(row.get("label", ""), url),
    }


def select_recent_api_context(info, recent_contexts):
    if not recent_contexts:
        return None
    api_name = (info.get("api_name") or "").strip()
    if api_name and api_name != "unknown":
        for ctx in reversed(recent_contexts):
            if ctx.get("api_name") == api_name or api_name in (ctx.get("url") or ""):
                return ctx
    return recent_contexts[-1]


def apply_api_context(info, ctx):
    if not ctx:
        return
    if not info.get("url") or info.get("url") == "null":
        info["url"] = ctx.get("url", "")
    if not info.get("method"):
        info["method"] = ctx.get("method", "")
    current_req = info.get("request_data") or ""
    if not current_req or looks_like_non_http_sampler_data(current_req):
        info["request_data"] = ctx.get("request_data", "")
    current_resp = info.get("response_data") or ""
    if not current_resp or looks_like_compare_report(current_resp) or not looks_like_json_payload(current_resp):
        info["response_data"] = ctx.get("response_data", "")
    if info.get("api_name") in ("", "unknown") or info.get("api_name") == info.get("label"):
        info["api_name"] = ctx.get("api_name") or info.get("api_name")


def infer_expected_source(message):
    if "DB/辅助接口/API" in message:
        return "DB/辅助接口"
    if re.search(r"\bCSV(?:\([^)]*\))?\s*=", message, re.I):
        return "CSV"
    if re.search(r"API\s+vs\s+DB|DB/API|数据库", message, re.I):
        return "DB"
    if re.search(r"API\s+vs\s+CSV|CSV/API", message, re.I):
        return "CSV"
    return "预期"


def split_pipe_segments(line):
    return [part.strip() for part in re.split(r"\s*\|\s*", line.strip()) if part.strip()]


def parse_bracket_value_compare_line(raw, status):
    """
    Parse lines like:
      字段 [name] 字符串不一致：API=[xxx]，CSV=[yyy]
      字段 [timestamp] 数值不一致: API=[0], DB=[1781152961]
    """
    m = re.search(
        r"字段\s*\[(?P<field>[^\]]+)\]\s*"
        r"(?P<reason>[^:：\n]*?不一致)\s*[:：]\s*"
        r"API\s*=\s*\[(?P<api>.*?)\]\s*[,，]\s*"
        r"(?P<expected_key>DB|CSV)(?:\((?P<expected_field>[^)]*)\))?\s*=\s*\[(?P<expected>.*?)\]\s*$",
        raw,
        re.I | re.S,
    )
    if not m:
        return None
    expected_key = (m.group("expected_key") or "").upper()
    expected_field = (m.group("expected_field") or "").strip()
    field = (m.group("field") or "").strip()
    source = expected_field or expected_key
    return {
        "field": field,
        "expected_value": (m.group("expected") or "").strip(),
        "actual_value": (m.group("api") or "").strip(),
        "aux_value": "",
        "db_field": source,
        "reason": (m.group("reason") or "字段不一致").strip(),
        "record_id": "",
        "status": status,
        "raw": raw,
    }


def parse_compare_line(line, status):
    """
    Parse field comparison lines, including examples:
      - [不同] | position=69936 | field=source | DB=mt4 | API=mt5 | DB字段=source
      失败 | 订单号=73193 | 字段=position_id | API=<empty> | DB=... | DB字段=position_id
      [失败] contract_size: API=0.0 | DB=100000.0 | DB字段=contract_size
      - [不同] | 登录号=4050 | field=equity | DB(equity/EQUITY/Equity)=45164.93 | get_user_info.Equity=45164.84 | API=45189
    """
    raw = line.strip().strip('"')
    if not raw:
        return None

    bracket_item = parse_bracket_value_compare_line(raw, status)
    if bracket_item:
        return bracket_item

    item = {
        "field": "",
        "expected_value": "",
        "actual_value": "",
        "aux_value": "",
        "db_field": "",
        "reason": "",
        "record_id": "",
        "status": status,
        "raw": raw,
    }

    rights_compare = re.search(
        r"\b(?P<api>change_user_rights|modify_user_info)\s+修改权限成\s+"
        r"(?P<changed_field>[A-Za-z_][A-Za-z0-9_]*)[：:]\s*(?P<changed_value>.*?)\s+"
        r"预期结果结果[：:]\s*get_user_info\s+响应信息"
        r"(?P<expected_field>[A-Za-z_][A-Za-z0-9_]*)[：:]\s*(?P<expected_value>.*?)\s+"
        r"实际结果[：:]\s*get_user_info\s+响应信息"
        r"(?P<actual_field>[A-Za-z_][A-Za-z0-9_]*)[：:]\s*(?P<actual_value>.*?)(?:\s*\||$)",
        raw,
        re.S,
    )
    if rights_compare:
        item["field"] = rights_compare.group("expected_field").strip() or rights_compare.group("changed_field").strip()
        item["expected_value"] = rights_compare.group("expected_value").strip()
        item["actual_value"] = rights_compare.group("actual_value").strip()
        item["db_field"] = f'get_user_info.{item["field"]}'
        item["aux_value"] = (
            f'{rights_compare.group("api").strip()} 修改'
            f'{rights_compare.group("changed_field").strip()}='
            f'{rights_compare.group("changed_value").strip()}'
        )
        item["reason"] = f'{rights_compare.group("api").strip()} 修改权限后，get_user_info 返回 {item["field"]} 与预期不一致'

    # Field name in "[失败] contract_size: API=..." style.
    prefix_field = re.search(r"^\s*-?\s*\[(?:不同|失败|跳过)\]\s*([^:|]+?)\s*:", raw)
    if prefix_field:
        item["field"] = prefix_field.group(1).strip()

    # Field name in "field=xxx" / "字段=xxx" style.
    m = re.search(r"(?:^|\|)\s*(?:field|字段)\s*=\s*([^|]+)", raw)
    if m:
        item["field"] = m.group(1).strip()

    # Record identifier.
    m = re.search(r"(?:position|订单号|登录号|login|orderId|order|ticket|matchedOrderId|group|currency)\s*=\s*([^|]+)", raw)
    if m:
        item["record_id"] = m.group(1).strip()

    aux_values = []
    if item["aux_value"]:
        aux_values.append(item["aux_value"])
    for segment in split_pipe_segments(raw):
        # Segment may be "[失败] contract_size: API=0.0".
        # Only split the colon when it appears before the first "=", otherwise
        # time values such as "2026-06-16 14:48:28" would be damaged.
        colon_pos = segment.find(":")
        equal_pos = segment.find("=")
        if colon_pos != -1 and equal_pos != -1 and colon_pos < equal_pos and re.search(r"\b(?:API|DB(?:\([^)]*\))?)=", segment):
            before, after = segment.split(":", 1)
            if not item["field"]:
                mf = re.search(r"\[(?:不同|失败|跳过)\]\s*(.+)$", before.strip())
                if mf:
                    item["field"] = mf.group(1).strip()
            segment = after.strip()

        if "=" not in segment:
            continue
        key, value = segment.split("=", 1)
        key = key.strip()
        value = value.strip()

        if key in ("field", "字段"):
            item["field"] = value
        elif key in ("position", "订单号", "登录号", "login", "order", "ticket", "matchedOrderId", "group", "currency"):
            item["record_id"] = value
        elif key == "API":
            item["actual_value"] = value
        elif key == "DB" or key.startswith("DB("):
            item["expected_value"] = value
            m_db = re.match(r"DB\(([^)]*)\)", key)
            if m_db and not item["db_field"]:
                item["db_field"] = m_db.group(1).strip()
        elif key == "CSV" or key.startswith("CSV("):
            item["expected_value"] = value
            m_csv = re.match(r"CSV\(([^)]*)\)", key)
            if m_csv and not item["db_field"]:
                item["db_field"] = m_csv.group(1).strip()
        elif key in ("DB字段", "DBField", "DB字段名", "CSV字段", "CSVField", "CSV字段名"):
            item["db_field"] = value
        elif key in ("reason", "原因"):
            item["reason"] = value
        elif key.startswith("get_user_info.") or key.startswith("get_position_ex."):
            aux_values.append(f"{key}={value}")
        elif key in ("字段来源", "fieldSource", "matched"):
            if not item["db_field"]:
                m_db = re.search(r"DB\(([^)]*)\)", value)
                if m_db:
                    item["db_field"] = m_db.group(1).strip()

    if not item["db_field"]:
        m_db = re.search(r"DB\(([^)]*)\)", raw)
        if m_db:
            item["db_field"] = m_db.group(1).strip()

    item["aux_value"] = " ; ".join(aux_values)

    if not item["field"]:
        return None
    return item


def dedupe_compare_items(items):
    result = []
    seen = set()
    for item in items:
        key = (
            normalize_compare_text(item.get("field")).lower(),
            normalize_compare_text(item.get("record_id")).lower(),
            normalize_compare_text(item.get("actual_value")),
            normalize_compare_text(item.get("expected_value")),
            normalize_compare_text(item.get("db_field")).lower(),
            normalize_compare_text(item.get("reason")),
        )
        if key in seen:
            continue
        seen.add(key)
        result.append(item)
    return result


def collect_compare_lines(message, kind):
    lines = []
    for raw_line in message.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if kind == "fail":
            if (
                line.startswith("失败 |")
                or line.startswith("对比失败 |")
                or re.match(r"^-?\s*\[(?:不同|失败)\]", line)
                or " [失败] " in line
                or (
                    line.startswith("字段 [")
                    and "不一致" in line
                    and re.search(r"\bAPI\s*=", line, re.I)
                    and re.search(r"\b(?:DB|CSV)(?:\([^)]*\))?\s*=", line, re.I)
                )
            ):
                lines.append(line)
        elif kind == "skip":
            if (
                line.startswith("跳过 |")
                or re.match(r"^-?\s*\[跳过\]", line)
                or " [跳过] " in line
            ):
                lines.append(line)
    return lines


def extract_counts(message):
    return {
        "total_fields": int(first_match([r"接口对比字段总数量[=：]\s*(\d+)", r"接口字段总数[=：]\s*(\d+)"], message, "0") or 0),
        "passed_count": int(first_match([r"相同字段数量[=：]\s*(\d+)", r"成功字段数量[=：]\s*(\d+)", r"成功数[=：]\s*(\d+)"], message, "0") or 0),
        "failed_count": int(first_match([r"不同字段数量[=：]\s*(\d+)", r"失败字段数量[=：]\s*(\d+)", r"失败数[=：]\s*(\d+)", r"共\s*(\d+)\s*处不一致"], message, "0") or 0),
        "skipped_count": int(first_match([r"跳过字段(?:（[^）]*）)?数量[=：]\s*(\d+)", r"跳过字段数量[=：]\s*(\d+)", r"跳过数[=：]\s*(\d+)"], message, "0") or 0),
    }


def empty_detail(row, message):
    label = row.get("label", "")
    url = row.get("URL", "")
    return {
        "label": label,
        "thread_name": row.get("threadName", ""),
        "url": url,
        "elapsed": row.get("elapsed", ""),
        "time_stamp": row.get("timeStamp", ""),
        "response_code": row.get("responseCode", ""),
        "response_message": row.get("responseMessage", ""),
        "response_data": response_data_from_row(row),
        "method": request_method_from_row(row),
        "request_data": request_data_from_row(row),
        "api_name": extract_api_name(label, url),
        "instance_id": "",
        "login": "",
        "table": "",
        "order_id": "",
        "expected_source": infer_expected_source(message),
        "total_fields": 0,
        "passed_count": 0,
        "failed_count": 0,
        "skipped_count": 0,
        "failed_fields": [],
        "skipped_fields": [],
        "api_issues": [],
        "error_type": "",
        "failure_message": message,
    }


def enrich_metadata(info):
    msg = info["failure_message"] or ""
    url = info["url"] or ""
    api_from_message = first_match([r"\b(change_user_rights|modify_user_info)\b"], msg)
    if api_from_message:
        info["api_name"] = api_from_message
    info["instance_id"] = first_match([r"instanceid=([^|\s,]+)"], msg) or query_value(url, "instanceid")
    info["login"] = (
        first_match([r"本次对比登录号=([^|\s,]+)", r"login=([^|\s,]+)", r"登录号=([^|\s,]+)"], msg)
        or query_value(url, "login")
        or query_value(url, "login__in")
    )
    info["table"] = first_match([r"table=([^|\s,]+)"], msg)
    info["order_id"] = first_match([r"matchedOrderId=([^|\s,]+)", r"匹配订单号=([^|\s,]+)"], msg)
    info.update(extract_counts(msg))


def is_empty_jmeter_sub_result(row):
    """Skip synthetic JMeter sub-results that do not carry the real failureMessage."""
    label = row.get("label", "") or ""
    if not re.search(r"-\d+$", label):
        return False
    if (row.get("failureMessage") or "").strip():
        return False
    url = (row.get("URL") or "").strip().lower()
    return url in ("", "null")


def extract_failures(rows):
    """Return only the meaningful failed samples, with structured comparison details."""
    results = []
    recent_api_contexts = []
    for row in rows:
        own_api_context = api_context_from_row(row)
        if str(row.get("success", "")).strip().lower() != "false":
            if own_api_context:
                recent_api_contexts.append(own_api_context)
                recent_api_contexts = recent_api_contexts[-12:]
            continue

        label = row.get("label", "")
        response_data = response_data_from_row(row)
        message = row.get("failureMessage", "") or ""
        if not message and looks_like_compare_report(response_data):
            message = response_data
        if not message:
            message = row.get("responseMessage", "") or ""

        # Transaction rows only say "number of failing samples"; child sampler has the real detail.
        if "事务" in label:
            if own_api_context:
                recent_api_contexts.append(own_api_context)
                recent_api_contexts = recent_api_contexts[-12:]
            continue

        # JMeter sub-results frequently end with "-0" and only repeat
        # "FAIL/compare failed"; keep real samplers such as "...断言-355"
        # when they carry failureMessage/URL details.
        if is_empty_jmeter_sub_result(row):
            if own_api_context:
                recent_api_contexts.append(own_api_context)
                recent_api_contexts = recent_api_contexts[-12:]
            continue

        info = empty_detail(row, message)
        enrich_metadata(info)
        apply_api_context(info, own_api_context or select_recent_api_context(info, recent_api_contexts))
        info["api_issues"] = detect_api_response_issues(info)

        fail_lines = collect_compare_lines(message, "fail")
        skip_lines = collect_compare_lines(message, "skip")
        info["failed_fields"] = dedupe_compare_items(
            [x for x in (parse_compare_line(line, "fail") for line in fail_lines) if x]
        )
        info["skipped_fields"] = dedupe_compare_items(
            [x for x in (parse_compare_line(line, "skip") for line in skip_lines) if x]
        )
        if info["failed_fields"] and not info["failed_count"]:
            info["failed_count"] = len(info["failed_fields"])
        if info["skipped_fields"] and not info["skipped_count"]:
            info["skipped_count"] = len(info["skipped_fields"])

        if any(x.get("type") == "API_ERROR" for x in info["api_issues"]):
            info["error_type"] = "API_ERROR"
        elif any(x.get("type") == "API_EMPTY_DATA" for x in info["api_issues"]):
            info["error_type"] = "API_EMPTY_DATA"
        elif info["failed_fields"] or info["skipped_fields"] or "对比结果" in message or "字段对比结果" in message:
            info["error_type"] = "COMPARE_FAIL"
        elif "数据库未返回" in message:
            info["error_type"] = "DB_NO_DATA"
        elif row.get("responseCode") and row.get("responseCode") not in ("200", "OK"):
            info["error_type"] = "HTTP_ERROR"
        else:
            info["error_type"] = "ASSERTION_FAIL"

        results.append(info)
        if own_api_context:
            recent_api_contexts.append(own_api_context)
            recent_api_contexts = recent_api_contexts[-12:]

    return results


def sample_time(ms):
    try:
        if not ms:
            return ""
        return datetime.fromtimestamp(int(ms) / 1000).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return ""


def is_transaction_row(row):
    label = row.get("label", "") or ""
    response_message = row.get("responseMessage", "") or ""
    return "事务" in label or "Number of samples in transaction" in response_message


def summarize(details, all_rows):
    raw_failed = [r for r in all_rows if str(r.get("success", "")).strip().lower() == "false"]
    transaction_rows = [r for r in all_rows if is_transaction_row(r)]
    failed_transaction_rows = [
        r for r in transaction_rows
        if str(r.get("success", "")).strip().lower() == "false"
    ]
    transaction_success = max(len(transaction_rows) - len(failed_transaction_rows), 0)
    fail_fields = []
    skip_fields = []
    for d in details:
        for f in d["failed_fields"]:
            fail_fields.append({**f, "api_name": d["api_name"], "label": d["label"]})
        for f in d["skipped_fields"]:
            skip_fields.append({**f, "api_name": d["api_name"], "label": d["label"]})

    api_groups = {}
    api_order = []
    for d in details:
        api_name = d["api_name"]
        if api_name not in api_groups:
            api_groups[api_name] = []
            api_order.append(api_name)
        api_groups[api_name].append(d)

    field_freq = {}
    for item in fail_fields:
        field_freq[item["field"]] = field_freq.get(item["field"], 0) + 1
    unique_fail_field_keys = {
        (item.get("api_name", ""), item.get("field", "").strip().lower())
        for item in fail_fields
        if item.get("field")
    }

    return {
        "total_rows": len(all_rows),
        "raw_failed": len(raw_failed),
        "transaction_total": len(transaction_rows),
        "transaction_failed": len(failed_transaction_rows),
        "transaction_success": transaction_success,
        "details": len(details),
        "fail_fields": fail_fields,
        "fail_field_total": len(fail_fields),
        "fail_field_unique_count": len(unique_fail_field_keys),
        "skip_fields": skip_fields,
        "api_groups": api_groups,
        "api_order": api_order,
        "field_freq": sorted(field_freq.items(), key=lambda x: (-x[1], x[0])),
    }


def rate_text(success, total):
    if not total:
        return "N/A"
    return f"{success / total * 100:.2f}%"


def iter_api_groups(summary):
    api_groups = summary["api_groups"]
    for api in summary.get("api_order") or list(api_groups.keys()):
        if api in api_groups:
            yield api, api_groups[api]


def css_block():
    return """
<style>
:root {
  --bg: #f6f7fb;
  --panel: #ffffff;
  --panel-2: #f9fafb;
  --border: #d9dde7;
  --text: #172033;
  --muted: #667085;
  --red: #d92d20;
  --red-bg: #fff1f0;
  --red-border: #f6b7ad;
  --green: #11845b;
  --green-bg: #ecfdf3;
  --yellow: #b54708;
  --yellow-bg: #fffaeb;
  --blue: #175cd3;
  --blue-bg: #eff8ff;
  --ink: #111827;
}
* { box-sizing: border-box; }
body {
  margin: 0;
  padding: 24px;
  background: var(--bg);
  color: var(--text);
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Microsoft YaHei", Arial, sans-serif;
  line-height: 1.55;
}
.wrap { max-width: 1480px; margin: 0 auto; }
.top {
  background: var(--panel);
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 20px 22px;
  margin-bottom: 16px;
  text-align: center;
}
.top h1 { margin: 0 0 6px; font-size: 24px; color: var(--ink); letter-spacing: 0; text-align: center; }
.top .sub { color: var(--muted); font-size: 13px; text-align: center; }
.stats {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
  gap: 12px;
  margin-bottom: 16px;
}
.stats.primary {
  grid-template-columns: repeat(4, minmax(0, 1fr));
}
.stats.secondary {
  grid-template-columns: repeat(3, minmax(0, 1fr));
}
.stat {
  background: var(--panel);
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 14px 16px;
}
.stat .num { font-size: 28px; font-weight: 750; line-height: 1.1; }
.stat .name { color: var(--muted); font-size: 12px; margin-top: 5px; }
.red { color: var(--red); }
.green { color: var(--green); }
.blue { color: var(--blue); }
.yellow { color: var(--yellow); }
.section {
  background: var(--panel);
  border: 1px solid var(--border);
  border-radius: 8px;
  margin-bottom: 16px;
  overflow: hidden;
}
.section-title {
  padding: 12px 16px;
  border-bottom: 1px solid var(--border);
  background: var(--panel-2);
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
}
.section-title h2 { margin: 0; font-size: 16px; color: var(--ink); }
.api-group {
  margin-top: 10px;
}
.api-group > summary.section-title {
  cursor: pointer;
  color: var(--ink);
  font-size: 16px;
  font-weight: 650;
  list-style: none;
  border-bottom: 0;
}
.api-group[open] > summary.section-title {
  border-bottom: 1px solid var(--border);
}
.api-group > summary.section-title::-webkit-details-marker {
  display: none;
}
.api-summary-title {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  min-width: 0;
}
.api-summary-title::before {
  content: ">";
  color: var(--red);
  font-size: 13px;
  line-height: 1;
}
.api-group[open] .api-summary-title::before {
  content: "v";
}
.api-summary-meta {
  display: inline-flex;
  flex-wrap: wrap;
  align-items: center;
  justify-content: flex-end;
  gap: 8px;
}
.api-group-body {
  background: var(--panel);
}
.api-field-summary {
  margin: 12px 16px 4px;
  padding: 10px 12px;
  border: 1px solid var(--red-border);
  background: var(--red-bg);
  border-radius: 7px;
}
.api-field-summary-title {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 10px;
  margin-bottom: 8px;
  color: var(--red);
  font-size: 13px;
  font-weight: 750;
}
.api-field-list {
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
}
.field-chip {
  display: inline-flex;
  align-items: center;
  gap: 5px;
  padding: 3px 8px;
  border: 1px solid var(--red-border);
  border-radius: 999px;
  background: #fff;
  color: var(--ink);
  font-family: Consolas, "Cascadia Code", monospace;
  font-size: 12px;
}
.field-chip-count {
  color: var(--red);
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Microsoft YaHei", Arial, sans-serif;
  font-weight: 750;
}
.api-field-reason,
.overview-reason {
  margin-top: 8px;
  color: #7a2e0e;
  font-size: 12px;
  line-height: 1.5;
}
.overview-reason {
  margin-top: 0;
}
.pill {
  display: inline-flex;
  align-items: center;
  padding: 2px 8px;
  border-radius: 999px;
  border: 1px solid var(--border);
  font-size: 12px;
  color: var(--muted);
  background: #fff;
  white-space: nowrap;
}
.pill.fail { color: var(--red); background: var(--red-bg); border-color: var(--red-border); }
.pill.skip { color: var(--yellow); background: var(--yellow-bg); border-color: #f6d7a8; }
.pill.ok { color: var(--green); background: var(--green-bg); border-color: #a7e3c1; }
.table { width: 100%; border-collapse: collapse; font-size: 13px; }
.table th {
  padding: 9px 10px;
  text-align: left;
  color: var(--muted);
  background: var(--panel-2);
  border-bottom: 1px solid var(--border);
  font-weight: 650;
  white-space: nowrap;
}
.table td {
  padding: 9px 10px;
  border-bottom: 1px solid #edf0f5;
  vertical-align: top;
}
.table tr:last-child td { border-bottom: 0; }
.sample {
  border-top: 1px solid var(--border);
  padding: 15px 16px 16px;
}
.sample:first-child { border-top: 0; }
.sample-head {
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
  align-items: center;
  margin-bottom: 8px;
}
.sample-title { font-size: 15px; font-weight: 700; color: var(--ink); }
.meta {
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
  margin: 7px 0 8px;
  color: var(--muted);
  font-size: 12px;
}
.meta span {
  background: var(--panel-2);
  border: 1px solid var(--border);
  padding: 2px 7px;
  border-radius: 5px;
}
.url {
  font-family: Consolas, "Cascadia Code", monospace;
  font-size: 12px;
  background: #f3f6fb;
  border: 1px solid var(--border);
  border-radius: 6px;
  padding: 7px 9px;
  word-break: break-all;
  color: #344054;
  margin: 7px 0 10px;
}
.params-block {
  margin: 8px 0 10px;
  border: 1px solid var(--border);
  border-radius: 7px;
  overflow: hidden;
  background: #fff;
}
.params-block summary {
  cursor: pointer;
  list-style: none;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 10px;
  padding: 7px 10px;
  background: #f8fafc;
  color: #344054;
  border-bottom: 1px solid var(--border);
}
.params-block:not([open]) summary {
  border-bottom: 0;
}
.params-block summary::-webkit-details-marker {
  display: none;
}
.params-summary-main {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  font-weight: 750;
}
.params-summary-main::before {
  content: ">";
  color: var(--blue);
  font-size: 13px;
  line-height: 1;
}
.params-block[open] .params-summary-main::before {
  content: "v";
}
.params-summary-sub {
  color: var(--muted);
  font-size: 12px;
}
.params-table {
  width: 100%;
  border-collapse: collapse;
  font-size: 12px;
}
.params-table th,
.params-table td {
  padding: 6px 9px;
  border-bottom: 1px solid #edf0f5;
  text-align: left;
  vertical-align: top;
}
.params-table th {
  width: 190px;
  color: var(--muted);
  background: #fff;
}
.params-table tr:last-child td {
  border-bottom: 0;
}
.params-table code {
  font-family: Consolas, "Cascadia Code", monospace;
  color: var(--ink);
  word-break: break-all;
}
.request-body-title {
  padding: 8px 10px 0;
  color: var(--muted);
  font-size: 12px;
  font-weight: 750;
}
.request-body {
  max-height: 300px;
  overflow: auto;
  margin: 6px 10px 10px;
  padding: 10px 12px;
  background: #f8fafc;
  border: 1px solid #edf0f5;
  border-radius: 6px;
  color: #344054;
  white-space: pre;
  font-family: Consolas, "Cascadia Code", monospace;
  font-size: 12px;
}
.response-block {
  margin: 8px 0 10px;
  border: 1px solid var(--red-border);
  border-radius: 7px;
  background: #fffafa;
  overflow: hidden;
}
.response-block summary {
  cursor: pointer;
  list-style: none;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 10px;
  padding: 8px 10px;
  color: var(--red);
  background: var(--red-bg);
  border-bottom: 1px solid var(--red-border);
}
.response-block:not([open]) summary {
  border-bottom: 0;
}
.response-block summary::-webkit-details-marker {
  display: none;
}
.response-summary-main {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  font-weight: 750;
}
.response-summary-main::before {
  content: ">";
  color: var(--red);
  font-size: 13px;
  line-height: 1;
}
.response-block[open] .response-summary-main::before {
  content: "v";
}
.response-summary-sub {
  color: #7a2e0e;
  font-size: 12px;
}
.response-fail-tags {
  display: flex;
  flex-wrap: wrap;
  gap: 6px;
  padding: 8px 10px;
  border-bottom: 1px solid var(--red-border);
}
.response-fail-tag {
  display: inline-flex;
  gap: 4px;
  align-items: center;
  padding: 2px 7px;
  border: 1px solid var(--red-border);
  border-radius: 999px;
  background: #fff;
  color: var(--red);
  font-family: Consolas, "Cascadia Code", monospace;
  font-size: 12px;
  font-weight: 700;
}
.response-body {
  max-height: 420px;
  overflow: auto;
  margin: 0;
  padding: 10px 12px;
  background: #fff;
  color: #344054;
  white-space: pre-wrap;
  word-break: break-word;
  font-family: Consolas, "Cascadia Code", monospace;
  font-size: 12px;
}
.json-body {
  white-space: pre;
}
.response-format-note {
  padding: 7px 10px;
  border-bottom: 1px solid var(--red-border);
  background: #fff;
  color: #7a2e0e;
  font-size: 12px;
}
.response-empty {
  padding: 10px 12px;
  color: #7a2e0e;
  background: #fff;
  font-size: 12px;
}
.resp-highlight {
  background: #ffe4e0;
  color: var(--red);
  border: 1px solid var(--red-border);
  border-radius: 4px;
  padding: 0 3px;
  font-weight: 750;
}
.resp-json-target-record {
  display: inline;
  background: rgba(255, 228, 224, 0.32);
}
.resp-json-failed-line {
  background: #ffe4e0;
  border: 1px solid var(--red-border);
  border-radius: 4px;
  padding: 1px 3px;
}
.resp-json-field-highlight,
.resp-json-value-highlight {
  background: #ffe4e0;
  color: var(--red);
  border: 1px solid var(--red-border);
  border-radius: 4px;
  padding: 0 3px;
  font-weight: 800;
}
.json-punc {
  color: #475467;
}
.compare-table {
  width: 100%;
  border-collapse: collapse;
  font-size: 12px;
  margin-top: 8px;
  border: 1px solid var(--border);
  border-radius: 6px;
  overflow: hidden;
}
.compare-table th {
  background: #f8fafc;
  color: var(--muted);
  text-align: left;
  padding: 8px;
  border-bottom: 1px solid var(--border);
}
.compare-table td {
  padding: 8px;
  border-bottom: 1px solid #edf0f5;
  vertical-align: top;
  word-break: break-word;
}
.compare-table tr:last-child td { border-bottom: 0; }
.compare-table tr.fail-row { background: var(--red-bg); }
.compare-table tr.skip-row { background: var(--yellow-bg); }
.failure-block {
  margin-top: 10px;
  border: 1px solid var(--red-border);
  background: #fffafa;
  border-radius: 7px;
  overflow: hidden;
}
.failure-block-title {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 10px;
  padding: 8px 10px;
  color: var(--red);
  background: var(--red-bg);
  border-bottom: 1px solid var(--red-border);
  font-size: 13px;
  font-weight: 750;
}
.failure-block .compare-table {
  margin-top: 0;
  border: 0;
  border-radius: 0;
}
.failure-block .raw {
  margin: 0;
  border: 0;
  border-radius: 0;
}
.value {
  display: inline-block;
  max-width: 100%;
  padding: 2px 6px;
  border-radius: 5px;
  font-family: Consolas, "Cascadia Code", monospace;
  font-size: 12px;
  white-space: pre-wrap;
  word-break: break-word;
}
.expected { background: var(--blue-bg); color: #1849a9; }
.actual { background: var(--red-bg); color: #b42318; }
.aux { background: var(--green-bg); color: #067647; }
.raw {
  margin-top: 9px;
  padding: 10px 12px;
  border: 1px solid var(--border);
  background: #fbfcff;
  border-radius: 6px;
  white-space: pre-wrap;
  font-family: Consolas, "Cascadia Code", monospace;
  font-size: 12px;
  color: #344054;
  overflow-x: auto;
}
details { margin-top: 10px; }
summary { cursor: pointer; color: var(--yellow); font-size: 13px; font-weight: 650; }
details.api-group { margin-top: 0; }
.empty {
  padding: 28px;
  text-align: center;
  color: var(--green);
  font-weight: 700;
}
@media (max-width: 760px) {
  body { padding: 10px; }
  .top h1 { font-size: 20px; }
  .stats,
  .stats.primary,
  .stats.secondary { grid-template-columns: repeat(2, minmax(0, 1fr)); }
  .compare-table { font-size: 11px; }
}
</style>
"""


def first_non_empty_segment(text):
    for segment in re.split(r"\s*\|\s*|\r?\n", text or ""):
        value = segment.strip().strip("：:").strip()
        if value:
            return value
    return ""


def bracket_fields(text, limit=8):
    fields = []
    seen = set()
    for field in re.findall(r"字段\s*\[([^\]]+)\]", text or ""):
        value = field.strip()
        key = value.lower()
        if value and key not in seen:
            seen.add(key)
            fields.append(value)
        if len(fields) >= limit:
            break
    return fields


def summarize_detail_failure_reason(detail, limit=180):
    message = detail.get("failure_message") or ""
    response_message = detail.get("response_message") or ""
    response_code = str(detail.get("response_code") or "").strip()
    raw = (message or response_message or "").replace("\r", "\n")
    normalized = re.sub(r"\s+", " ", raw).strip(" |")

    api_issues = detail.get("api_issues") or []
    if api_issues:
        issue = api_issues[0]
        code_msg = []
        if issue.get("code"):
            code_msg.append(f"code={issue.get('code')}")
        if issue.get("msg"):
            code_msg.append(f"msg={issue.get('msg')}")
        suffix = f"（{', '.join(code_msg)}）" if code_msg else ""
        return shorten_text(f"{issue.get('title') or '接口失败'}：{issue.get('reason') or issue.get('actual_value')}{suffix}", limit)

    if not normalized and response_message:
        normalized = re.sub(r"\s+", " ", response_message).strip()

    m = re.search(r"Test failed:\s*text expected to contain\s*/(.+?)/", normalized)
    if m:
        expected = m.group(1).replace(r"\s*", "").replace("\\", "")
        return shorten_text(f"响应断言失败：响应未包含期望内容 {expected}", limit)

    if "接口没有返回可对比" in normalized:
        return shorten_text(first_non_empty_segment(normalized) or normalized, limit)

    mismatch_count = first_match([r"共\s*(\d+)\s*处不一致"], normalized)
    fields = bracket_fields(normalized)
    field_text = "、".join(fields)
    if fields and len(fields) >= 8:
        field_text += " 等"

    if "数据库变量" in normalized and "不存在" in normalized:
        count_text = f"共 {mismatch_count} 处；" if mismatch_count else ""
        if field_text:
            return shorten_text(f"数据库变量不存在，可能查询无结果或字段名不匹配；{count_text}字段：{field_text}", limit)
        return shorten_text("数据库变量不存在，可能查询无结果或字段名不匹配", limit)

    if "API对比失败" in normalized or mismatch_count:
        title = first_non_empty_segment(raw) or first_non_empty_segment(normalized)
        if title.startswith("【") and title.endswith("】"):
            title = title.strip("【】")
        if field_text:
            title = f"{title}；字段：{field_text}"
        return shorten_text(title, limit)

    if detail.get("error_type") == "HTTP_ERROR":
        reason = f"HTTP {response_code}：{response_message or first_non_empty_segment(normalized) or '请求失败'}"
        return shorten_text(reason, limit)

    if response_message and response_message not in ("OK", "null"):
        return shorten_text(response_message, limit)

    if normalized:
        return shorten_text(first_non_empty_segment(normalized) or normalized, limit)

    return "未解析到字段明细，请展开查看原始 failureMessage"


def summarize_no_field_reasons(items, max_reasons=3, limit=240):
    reason_counts = {}
    reason_order = []
    for detail in items:
        if detail.get("failed_fields"):
            continue
        reason = summarize_detail_failure_reason(detail, limit=limit)
        if not reason:
            continue
        if reason not in reason_counts:
            reason_order.append(reason)
            reason_counts[reason] = 0
        reason_counts[reason] += 1

    if not reason_order:
        return ""

    parts = []
    for reason in reason_order[:max_reasons]:
        count = reason_counts[reason]
        parts.append(f"{reason}（{count}条）" if count > 1 else reason)
    if len(reason_order) > max_reasons:
        parts.append(f"另有 {len(reason_order) - max_reasons} 类原因")
    return "；".join(parts)


def render_overview(summary):
    rows = []
    for api, items in iter_api_groups(summary):
        fail_count = sum(len(d["failed_fields"]) for d in items)
        field_counts = {}
        for item in summary["fail_fields"]:
            if item["api_name"] == api:
                field_counts[item["field"]] = field_counts.get(item["field"], 0) + 1
        top_fields = sorted(field_counts.items(), key=lambda x: (-x[1], x[0]))[:8]
        no_field_reason = summarize_no_field_reasons(items)
        if top_fields:
            top_field_html = h("、".join(f"{field}({count})" for field, count in top_fields))
            if no_field_reason:
                top_field_html += f'<div class="overview-reason">另有无字段失败：{h(no_field_reason)}</div>'
        else:
            top_field_html = (
                f'<div class="overview-reason">失败原因：{h(no_field_reason)}</div>'
                if no_field_reason
                else '<div class="overview-reason">失败原因：未解析到字段明细，请展开查看原始 failureMessage</div>'
            )
        rows.append(
            "<tr>"
            f"<td><strong>{h(api)}</strong></td>"
            f"<td>{len(items)}</td>"
            f"<td class=\"red\">{fail_count}</td>"
            f"<td>{top_field_html}</td>"
            "</tr>"
        )
    return (
        '<div class="section"><div class="section-title"><h2>接口失败统计</h2>'
        f'<span class="pill fail">{len(summary["api_groups"])} 个接口涉及失败</span></div>'
        '<table class="table"><thead><tr><th>接口</th><th>失败样本数</th><th>失败字段明细数</th><th>高频失败字段</th></tr></thead><tbody>'
        + "".join(rows)
        + "</tbody></table></div>"
    )


def render_compare_rows(fields, expected_source, row_class):
    rows = []
    for item in fields:
        aux = item.get("aux_value") or ""
        rows.append(
            f'<tr class="{row_class}">'
            f'<td><strong>{h(item.get("field"))}</strong></td>'
            f'<td>{h(item.get("record_id"))}</td>'
            f'<td><span class="value actual">{h(item.get("actual_value"))}</span></td>'
            f'<td><span class="value expected">{h(item.get("expected_value"))}</span></td>'
            f'<td>{h(item.get("db_field"))}</td>'
            f'<td><span class="value aux">{h(aux)}</span></td>'
            f'<td>{h(item.get("reason"))}</td>'
            "</tr>"
        )
    return (
        '<table class="compare-table"><thead><tr>'
        '<th>字段</th><th>订单/登录</th>'
        '<th>API 实际值</th>'
        f'<th>{h(expected_source)} 预期值</th>'
        '<th>DB/CSV字段来源</th><th>辅助接口值</th><th>原因</th>'
        '</tr></thead><tbody>'
        + "".join(rows)
        + "</tbody></table>"
    )


def render_api_issue_rows(issues):
    rows = []
    for item in issues:
        code_msg = []
        if item.get("code"):
            code_msg.append(f"code={item.get('code')}")
        if item.get("msg"):
            code_msg.append(f"msg={item.get('msg')}")
        if item.get("meta") and item.get("meta") not in ("null", "{}", "[]", '""'):
            code_msg.append(f"meta={item.get('meta')}")
        rows.append(
            '<tr class="fail-row">'
            f'<td><strong>{h(item.get("title"))}</strong></td>'
            f'<td>{h(item.get("node"))}</td>'
            f'<td><span class="value actual">{h(item.get("actual_value"))}</span></td>'
            f'<td><span class="value expected">{h(item.get("expected_value"))}</span></td>'
            f'<td>{h(" | ".join(code_msg))}</td>'
            f'<td>{h(item.get("reason"))}</td>'
            '</tr>'
        )
    return (
        '<table class="compare-table"><thead><tr>'
        '<th>失败类型</th><th>响应节点</th><th>API 实际值</th><th>预期结果</th><th>接口 code/msg/meta</th><th>失败原因</th>'
        '</tr></thead><tbody>'
        + "".join(rows)
        + "</tbody></table>"
    )


def render_api_issue_block(detail):
    issues = detail.get("api_issues") or []
    if not issues:
        return ""
    issue_types = []
    for item in issues:
        title = item.get("title") or "接口失败"
        if title not in issue_types:
            issue_types.append(title)
    title_text = "、".join(issue_types)
    return (
        '<div class="failure-block">'
        '<div class="failure-block-title">'
        f'<span>失败明细: {h(title_text)}</span>'
        '<span>接口响应未满足可对比条件</span>'
        '</div>'
        + render_api_issue_rows(issues)
        + '</div>'
    )


RESPONSE_RECORD_ID_KEYS = (
    "order_ticket",
    "position_ticket",
    "deal_ticket",
    "ticket",
    "Ticket",
    "ticket_id",
    "TicketID",
    "order",
    "Order",
    "order_id",
    "orderId",
    "OrderID",
    "OrderTicket",
    "position",
    "Position",
    "position_id",
    "positionId",
    "PositionID",
    "PositionTicket",
    "deal",
    "Deal",
    "deal_id",
    "dealId",
    "DealID",
    "matchedOrderId",
    "login",
    "LOGIN",
    "Login",
    "group",
    "Group",
    "currency",
    "Currency",
)
RESPONSE_RECORD_ID_KEY_SET = {key.lower() for key in RESPONSE_RECORD_ID_KEYS}
STAT_SUFFIXES = ("stddev", "sum", "avg", "max", "min", "count")


def normalize_compare_text(value):
    if value is None:
        return ""
    return str(value).strip().strip('"').strip("'")


def normalized_failed_field_names(detail):
    return {
        normalize_compare_text(item.get("field")).lower()
        for item in detail.get("failed_fields", [])
        if normalize_compare_text(item.get("field"))
    }


def failed_field_actual_values(detail):
    values = {}
    for item in detail.get("failed_fields", []):
        field = normalize_compare_text(item.get("field")).lower()
        actual = normalize_compare_text(item.get("actual_value"))
        if not field or not actual:
            continue
        values.setdefault(field, []).append(actual)
    return values


def failed_field_path_actual_values(detail):
    values = {}
    for item in detail.get("failed_fields", []):
        field = normalize_compare_text(item.get("field")).lower()
        actual = normalize_compare_text(item.get("actual_value"))
        if not field or not actual:
            continue
        values.setdefault((field,), []).append(actual)
        for suffix in STAT_SUFFIXES:
            marker = "_" + suffix
            if field.endswith(marker) and len(field) > len(marker):
                base = field[: -len(marker)]
                values.setdefault((base, suffix), []).append(actual)
                values.setdefault(("stat", base, suffix), []).append(actual)
                break
    return values


def response_path_actuals(path, failed_paths):
    matched = []
    if not path:
        return matched
    lower_path = tuple(str(x).lower() for x in path)
    for suffix, actuals in failed_paths.items():
        if len(suffix) <= len(lower_path) and lower_path[-len(suffix):] == suffix:
            matched.extend(actuals)
    return matched


def response_path_is_failed_branch(path, failed_paths):
    if not path:
        return False
    lower_path = tuple(str(x).lower() for x in path)
    for suffix in failed_paths:
        if not suffix:
            continue
        for length in range(1, len(suffix)):
            if len(lower_path) >= length and lower_path[-length:] == suffix[:length]:
                return True
    return False


def actual_value_candidates(value):
    text = normalize_compare_text(value)
    if not text:
        return []
    if text == "<empty>":
        return ["", "null"]
    candidates = [text]
    for sep in ("->", "；", ";"):
        if sep in text:
            first = text.split(sep, 1)[0].strip()
            if first:
                candidates.append(first)
    result = []
    seen = set()
    for item in candidates:
        key = item.lower()
        if key not in seen:
            seen.add(key)
            result.append(item)
    return result


def numeric_equal(left, right):
    try:
        return float(left) == float(right)
    except Exception:
        return False


def response_value_matches_actual(response_value, actual_values):
    response_text = normalize_json_scalar(response_value)
    response_key = response_text.lower()
    for actual in actual_values:
        for candidate in actual_value_candidates(actual):
            candidate_text = normalize_compare_text(candidate)
            if response_key == candidate_text.lower():
                return True
            if numeric_equal(response_text, candidate_text):
                return True
    return False


def response_target_ids(detail):
    explicit = []
    for item in detail.get("failed_fields", []):
        record_id = normalize_compare_text(item.get("record_id"))
        if record_id:
            explicit.append(record_id)
    order_id = normalize_compare_text(detail.get("order_id"))
    if order_id:
        explicit.append(order_id)

    values = explicit
    if not values:
        login = normalize_compare_text(detail.get("login"))
        if login:
            values = [login]

    seen = set()
    result = []
    for value in values:
        key = value.lower()
        if key not in seen:
            seen.add(key)
            result.append(value)
    return result


def normalize_json_scalar(value):
    if value is None:
        return "null"
    if isinstance(value, bool):
        return "true" if value else "false"
    return str(value).strip()


def dict_record_values(obj):
    values = []
    if not isinstance(obj, dict):
        return values
    for key, value in obj.items():
        if str(key).lower() in RESPONSE_RECORD_ID_KEY_SET:
            values.append(normalize_json_scalar(value))
    return values


def json_value_by_key_ci(obj, wanted_key):
    if not isinstance(obj, dict):
        return None
    wanted = wanted_key.lower()
    for key, value in obj.items():
        if str(key).lower() == wanted:
            return value
    return None


def response_target_filters(detail):
    filters = []
    instance_id = normalize_compare_text(detail.get("instance_id"))
    if instance_id:
        filters.append(("instance_id", instance_id))
    return filters


def record_matches_target_filters(obj, target_filters):
    if not isinstance(obj, dict):
        return False
    for key, expected in target_filters:
        actual = json_value_by_key_ci(obj, key)
        if actual is not None and normalize_json_scalar(actual).lower() != expected.lower():
            return False
    return True


def is_target_json_record(obj, target_ids, target_filters=None):
    if not isinstance(obj, dict) or not target_ids:
        return False
    lower_targets = {str(x).lower() for x in target_ids if str(x).strip()}
    if not any(value.lower() in lower_targets for value in dict_record_values(obj)):
        return False
    return record_matches_target_filters(obj, target_filters or [])


def json_has_target_record(value, target_ids, target_filters=None):
    if not target_ids:
        return False
    if isinstance(value, dict):
        if is_target_json_record(value, target_ids, target_filters):
            return True
        return any(json_has_target_record(child, target_ids, target_filters) for child in value.values())
    if isinstance(value, list):
        return any(json_has_target_record(item, target_ids, target_filters) for item in value)
    return False


def parse_json_response(response_data):
    text = (response_data or "").strip()
    if not text or text[:1] not in ("{", "["):
        return None
    try:
        return json.loads(text)
    except Exception:
        return None


def json_scalar_literal(value):
    return json.dumps(value, ensure_ascii=False)


def render_json_scalar(value, highlight=False):
    literal = h(json_scalar_literal(value))
    if highlight:
        return f'<span class="resp-json-value-highlight">{literal}</span>'
    return literal


def render_json_response_value(
    value,
    failed_actuals,
    target_ids,
    target_filters=None,
    depth=0,
    parent_is_target=False,
    failed_paths=None,
    path=(),
):
    indent = "  " * depth
    child_indent = "  " * (depth + 1)
    highlight_without_target = not target_ids
    failed_paths = failed_paths or {}

    if isinstance(value, dict):
        current_is_target = parent_is_target or is_target_json_record(value, target_ids, target_filters)
        if not value:
            return '<span class="json-punc">{}</span>'

        rows = ['<span class="json-punc">{</span>']
        items = list(value.items())
        for index, (key, child) in enumerate(items):
            key_text = str(key)
            field_key = key_text.lower()
            child_path = path + (field_key,)
            path_actuals = response_path_actuals(child_path, failed_paths)
            actual_values = list(failed_actuals.get(field_key, [])) + path_actuals
            is_failed_key = field_key in failed_actuals or bool(path_actuals)
            highlight_value = (
                is_failed_key
                and not isinstance(child, (dict, list))
                and (current_is_target or highlight_without_target)
                and response_value_matches_actual(child, actual_values)
            )
            key_html = h(json.dumps(key_text, ensure_ascii=False))
            if highlight_value:
                key_html = f'<span class="resp-json-field-highlight">{key_html}</span>'
            child_html = render_json_response_value(
                child,
                failed_actuals,
                target_ids,
                target_filters,
                depth + 1,
                parent_is_target=current_is_target,
                failed_paths=failed_paths,
                path=child_path,
            )
            if highlight_value:
                child_html = render_json_scalar(child, highlight=True)
            comma = "," if index < len(items) - 1 else ""
            rows.append(f'\n{child_indent}<span>{key_html}: {child_html}</span>{comma}')
        rows.append(f'\n{indent}<span class="json-punc">}}</span>')
        html = "".join(rows)
        if current_is_target and target_ids:
            return f'<span class="resp-json-target-record">{html}</span>'
        return html

    if isinstance(value, list):
        if not value:
            return '<span class="json-punc">[]</span>'
        rows = ['<span class="json-punc">[</span>']
        for index, item in enumerate(value):
            item_html = render_json_response_value(
                item,
                failed_actuals,
                target_ids,
                target_filters,
                depth + 1,
                parent_is_target=parent_is_target,
                failed_paths=failed_paths,
                path=path,
            )
            comma = "," if index < len(value) - 1 else ""
            rows.append(f"\n{child_indent}{item_html}{comma}")
        rows.append(f'\n{indent}<span class="json-punc">]</span>')
        return "".join(rows)

    return render_json_scalar(value)


def response_highlight_tokens(detail):
    tokens = []
    seen = set()
    for item in detail.get("failed_fields", []):
        text = normalize_compare_text(item.get("actual_value"))
        if not text or len(text) > 120:
            continue
        if text.lower() in ("0", "0.0", "1", "1.0", "true", "false", "null", "<empty>", "not_found"):
            continue
        key = text.lower()
        if key not in seen:
            seen.add(key)
            tokens.append(text)
    return tokens


def highlight_response_text(text, tokens):
    if not text:
        return ""
    valid = [t for t in tokens if t]
    if not valid:
        return h(text)
    pattern = re.compile("|".join(re.escape(t) for t in sorted(valid, key=len, reverse=True)), re.I)
    pieces = []
    last = 0
    for match in pattern.finditer(text):
        if match.start() > last:
            pieces.append(h(text[last:match.start()]))
        pieces.append(f'<span class="resp-highlight">{h(match.group(0))}</span>')
        last = match.end()
    pieces.append(h(text[last:]))
    return "".join(pieces)


def render_response_fail_tags(failed_fields):
    if not failed_fields:
        return ""
    tags = []
    seen = set()
    for item in failed_fields:
        field = normalize_compare_text(item.get("field"))
        actual = normalize_compare_text(item.get("actual_value"))
        record_id = normalize_compare_text(item.get("record_id"))
        if not field and not actual:
            continue
        label = field if field else "actual"
        value = f"={actual}" if actual else ""
        record = f" / {record_id}" if record_id else ""
        key = (label.lower(), value.lower(), record.lower())
        if key in seen:
            continue
        seen.add(key)
        tags.append(f'<span class="response-fail-tag">{h(label)}{h(value)}{h(record)}</span>')
    if not tags:
        return ""
    return '<div class="response-fail-tags">' + "".join(tags) + "</div>"


def render_response_issue_tags(issues):
    if not issues:
        return ""
    tags = []
    seen = set()
    for item in issues:
        node = normalize_compare_text(item.get("node")) or "response"
        actual = normalize_compare_text(item.get("actual_value"))
        label = f"{node}={actual}" if actual else node
        key = label.lower()
        if key in seen:
            continue
        seen.add(key)
        tags.append(f'<span class="response-fail-tag">{h(label)}</span>')
    if not tags:
        return ""
    return '<div class="response-fail-tags">' + "".join(tags) + "</div>"


def render_response_content(detail, response_data):
    parsed = parse_json_response(response_data)
    if parsed is not None:
        failed_actuals = failed_field_actual_values(detail)
        failed_paths = failed_field_path_actual_values(detail)
        target_ids = response_target_ids(detail)
        target_filters = response_target_filters(detail)
        target_found = json_has_target_record(parsed, target_ids, target_filters)
        effective_target_ids = target_ids if target_found else []
        effective_filters = target_filters if target_found else []
        body = render_json_response_value(parsed, failed_actuals, effective_target_ids, effective_filters, failed_paths=failed_paths)
        if target_found:
            target_note = f"已按记录 {', '.join(target_ids)} 定位高亮"
        elif target_ids:
            target_note = f"响应中未找到记录 {', '.join(target_ids)}，已按失败字段名高亮"
        else:
            target_note = "未识别订单/登录号，按失败字段名高亮"
        return (
            '<div class="response-format-note">JSON 已格式化，失败明细中的 API 实际值已标红；'
            f'{h(target_note)}。</div>'
            f'<pre class="response-body json-body">{body}</pre>'
        )

    tokens = response_highlight_tokens(detail)
    body = highlight_response_text(response_data, tokens)
    return f'<pre class="response-body">{body}</pre>'


def render_response_block(detail):
    failed_fields = detail.get("failed_fields", [])
    response_data = detail.get("response_data") or ""
    tag_html = render_response_issue_tags(detail.get("api_issues") or []) + render_response_fail_tags(failed_fields)

    if response_data:
        size_text = f"{len(response_data)} 字符，点击展开"
        content = tag_html + render_response_content(detail, response_data)
    else:
        size_text = "JTL 未保存响应正文，点击展开"
        content = (
            tag_html
            + '<div class="response-empty">'
            '当前 JTL 未保存接口响应正文。需要在 JMeter 保存结果配置中开启 responseData/responseBody 后，报告会在这里显示完整响应，并自动用红色标记失败字段和 API 实际值。'
            '</div>'
        )

    return (
        '<details class="response-block">'
        '<summary>'
        '<span class="response-summary-main">接口响应</span>'
        f'<span class="response-summary-sub">{h(size_text)}</span>'
        '</summary>'
        + content
        + '</details>'
    )


def request_params_from_url(url):
    if not url or url == "null":
        return []
    try:
        parsed = urlparse(url)
        if not parsed.query:
            return []
        return [(key, value) for key, value in parse_qsl(parsed.query, keep_blank_values=True)]
    except Exception:
        return []


def compact_json_value(value):
    if isinstance(value, (dict, list)):
        return json.dumps(value, ensure_ascii=False, separators=(",", ":"))
    return normalize_json_scalar(value)


def flatten_json_params(value, prefix=""):
    params = []
    if isinstance(value, dict):
        for key, child in value.items():
            name = f"{prefix}.{key}" if prefix else str(key)
            if isinstance(child, dict):
                params.extend(flatten_json_params(child, name))
            elif isinstance(child, list):
                params.append((name, compact_json_value(child)))
            else:
                params.append((name, normalize_json_scalar(child)))
    elif isinstance(value, list):
        params.append((prefix or "body", compact_json_value(value)))
    elif prefix:
        params.append((prefix, normalize_json_scalar(value)))
    return params


def extract_request_body_text(request_data):
    text = (request_data or "").strip()
    if not text:
        return ""
    if looks_like_json_payload(text):
        return text
    marker = re.search(r"(?:POST|PUT|PATCH|DELETE)\s+data:\s*(.*)", text, re.I | re.S)
    if marker:
        body = marker.group(1).strip()
        body = re.split(r"\n\s*\[(?:no cookies|cookies|headers?)\]", body, 1, flags=re.I)[0].strip()
        return body
    marker = re.search(r"(?:Request Body|Body)\s*[:：]\s*(.*)", text, re.I | re.S)
    if marker:
        return marker.group(1).strip()
    return ""


def pretty_request_body(request_data):
    body = extract_request_body_text(request_data)
    if not body:
        return ""
    try:
        parsed = json.loads(body)
        return json.dumps(parsed, ensure_ascii=False, indent=2)
    except Exception:
        return body


def request_params_from_data(request_data):
    text = extract_request_body_text(request_data) or (request_data or "").strip()
    if not text:
        return []
    if looks_like_json_payload(text):
        try:
            parsed = json.loads(text)
            return flatten_json_params(parsed)
        except Exception:
            return [("body", text)]
    if "=" in text and "\n" not in text[:200]:
        try:
            parsed = parse_qsl(text, keep_blank_values=True)
            if parsed:
                return parsed
        except Exception:
            pass
    return [("body", text)]


def fallback_request_params(detail):
    candidates = [
        ("instanceid", detail.get("instance_id")),
        ("login", detail.get("login")),
        ("matchedOrderId", detail.get("order_id")),
        ("table", detail.get("table")),
    ]
    params = []
    for key, value in candidates:
        text = normalize_compare_text(value)
        if text:
            params.append((key, text))
    return params


def render_request_params(detail):
    method = (detail.get("method") or "").strip().upper()
    if method in ("GET", "HEAD"):
        return ""
    request_data = detail.get("request_data") or ""
    actual_params = request_params_from_url(detail.get("url"))
    if request_data and not looks_like_non_http_sampler_data(request_data):
        actual_params.extend(request_params_from_data(request_data))

    params = []
    seen_keys = set()
    for key, value in actual_params:
        key_text = normalize_compare_text(key)
        if not key_text:
            continue
        lower_key = key_text.lower()
        if lower_key in seen_keys:
            continue
        seen_keys.add(lower_key)
        params.append((key_text, value))

    from_fallback = False
    if not params:
        from_fallback = True
        for key, value in fallback_request_params(detail):
            key_text = normalize_compare_text(key)
            text = normalize_compare_text(value)
            if key_text and text:
                params.append((key_text, text))

    body_text = ""
    if request_data and not looks_like_non_http_sampler_data(request_data):
        body_text = pretty_request_body(request_data)

    if not params and not body_text:
        return ""
    rows = []
    for key, value in params:
        rows.append(
            "<tr>"
            f"<td>{h(key)}</td>"
            f"<td><code>{h(value)}</code></td>"
            "</tr>"
        )
    table_html = ""
    if rows:
        table_html = (
            '<table class="params-table"><thead><tr><th>参数</th><th>值</th></tr></thead><tbody>'
            + "".join(rows)
            + "</tbody></table>"
        )
    body_html = ""
    if body_text:
        body_html = (
            '<div class="request-body-title">完整 Request Body（按 JMeter 请求体格式化）</div>'
            f'<pre class="request-body">{h(body_text)}</pre>'
        )
    suffix = "，含完整 Request Body" if body_text else ""
    if from_fallback:
        suffix = "，JTL未保存真实请求参数，展示失败信息上下文" + suffix
    return (
        '<details class="params-block" open>'
        '<summary>'
        '<span class="params-summary-main">请求传参</span>'
        f'<span class="params-summary-sub">{len(params)} 个参数{h(suffix)}，点击收起/展开</span>'
        '</summary>'
        + table_html
        + body_html
        + '</details>'
    )


def render_api_failed_field_summary(items):
    field_counts = {}
    field_names = {}
    for detail in items:
        for item in detail.get("failed_fields", []):
            field = (item.get("field") or "").strip()
            if not field:
                continue
            key = field.lower()
            field_counts[key] = field_counts.get(key, 0) + 1
            field_names.setdefault(key, field)

    no_field_reason = summarize_no_field_reasons(items)

    if not field_counts:
        reason_html = (
            f'<div class="api-field-reason"><strong>失败原因：</strong>{h(no_field_reason)}</div>'
            if no_field_reason
            else '<div class="api-field-reason"><strong>失败原因：</strong>未解析到字段明细，请展开查看原始 failureMessage</div>'
        )
        return (
            '<div class="api-field-summary">'
            '<div class="api-field-summary-title">'
            '<span>对比不一致字段(去重)：0 个</span>'
            '<span>无字段失败明细，已补充失败原因</span>'
            '</div>'
            + reason_html
            + '</div>'
        )

    chips = []
    for key, count in sorted(field_counts.items(), key=lambda x: (-x[1], field_names[x[0]].lower())):
        chips.append(
            '<span class="field-chip">'
            f'<span>{h(field_names[key])}</span>'
            f'<span class="field-chip-count">{count}次</span>'
            '</span>'
        )

    return (
        '<div class="api-field-summary">'
        '<div class="api-field-summary-title">'
        f'<span>对比不一致字段(去重)：{len(field_counts)} 个</span>'
        f'<span>字段失败明细 {sum(field_counts.values())} 个</span>'
        '</div>'
        '<div class="api-field-list">'
        + "".join(chips)
        + '</div>'
        + (
            f'<div class="api-field-reason"><strong>另有无字段失败原因：</strong>{h(no_field_reason)}</div>'
            if no_field_reason
            else ""
        )
        + '</div>'
    )


def render_sample(detail, index):
    pieces = []
    skip_count = len(detail["skipped_fields"])
    badge_text = {
        "COMPARE_FAIL": "字段对比失败",
        "DB_NO_DATA": "数据库无数据",
        "API_ERROR": "接口业务失败",
        "API_EMPTY_DATA": "接口data为空",
        "HTTP_ERROR": "HTTP错误",
        "ASSERTION_FAIL": "断言失败",
    }.get(detail["error_type"], "失败")
    badge_class = "skip" if detail["error_type"] == "DB_NO_DATA" else "fail"

    pieces.append('<div class="sample">')
    pieces.append(
        '<div class="sample-head">'
        f'<span class="pill fail">#{index}</span>'
        f'<span class="sample-title">{h(detail["label"])}</span>'
        f'<span class="pill {badge_class}">{h(badge_text)}</span>'
        f'<span class="pill">{h(detail["api_name"])}</span>'
        '</div>'
    )

    meta = []
    if detail["instance_id"]:
        meta.append(f"实例: {h(detail['instance_id'])}")
    if detail["login"]:
        meta.append(f"登录号: {h(detail['login'])}")
    if detail["order_id"]:
        meta.append(f"匹配订单: {h(detail['order_id'])}")
    if detail["table"]:
        meta.append(f"DB表: {h(detail['table'])}")
    if detail["elapsed"]:
        meta.append(f"耗时: {h(detail['elapsed'])} ms")
    if detail["response_code"]:
        meta.append(f"响应码: {h(detail['response_code'])}")
    st = sample_time(detail["time_stamp"])
    if st:
        meta.append(f"时间: {h(st)}")
    if meta:
        pieces.append('<div class="meta">' + "".join(f"<span>{item}</span>" for item in meta) + "</div>")

    if detail["url"] and detail["url"] != "null":
        pieces.append(f'<div class="url">{h(detail["url"])}</div>')

    pieces.append(render_request_params(detail))

    if detail["total_fields"]:
        pieces.append(
            '<div class="meta">'
            f'<span>接口字段总数: {detail["total_fields"]}</span>'
            f'<span class="green">成功: {detail["passed_count"]}</span>'
            f'<span class="red">失败: {detail["failed_count"]}</span>'
            f'<span class="yellow">跳过: {detail["skipped_count"]}</span>'
            "</div>"
        )

    failure_message_rendered = False
    api_issue_html = render_api_issue_block(detail)
    if api_issue_html:
        pieces.append(api_issue_html)
        failure_message_rendered = True
        if detail["failure_message"]:
            pieces.append(f'<details><summary>查看原始 failureMessage</summary><div class="raw">{h(detail["failure_message"])}</div></details>')

    if detail["failed_fields"]:
        pieces.append(
            '<div class="failure-block">'
            '<div class="failure-block-title">'
            f'<span>失败明细: {len(detail["failed_fields"])} 个字段，已打印预期结果与实际结果</span>'
            f'<span>{h(detail["expected_source"])} vs API</span>'
            '</div>'
            + render_compare_rows(detail["failed_fields"], detail["expected_source"], "fail-row")
            + "</div>"
        )
    elif not detail.get("api_issues") and (detail["failed_count"] or detail["error_type"] == "COMPARE_FAIL"):
        failure_message_rendered = True
        pieces.append(
            '<div class="failure-block">'
            '<div class="failure-block-title">'
            '<span>失败明细: 未解析到结构化字段，直接打印原始失败内容</span>'
            '</div>'
            f'<div class="raw">{h(detail["failure_message"] or detail["response_message"] or "无 failureMessage")}</div>'
            '</div>'
        )

    if detail["skipped_fields"]:
        pieces.append(
            f'<details><summary>跳过字段明细: {skip_count} 个，展开查看 DB/CSV 与 API 数值</summary>'
            + render_compare_rows(detail["skipped_fields"], detail["expected_source"], "skip-row")
            + "</details>"
        )

    pieces.append(render_response_block(detail))

    if not detail["failed_fields"] and not detail["skipped_fields"] and not failure_message_rendered:
        pieces.append(f'<div class="raw">{h(detail["failure_message"] or detail["response_message"] or "无 failureMessage")}</div>')
    elif not failure_message_rendered:
        pieces.append(f'<details><summary>查看原始 failureMessage</summary><div class="raw">{h(detail["failure_message"])}</div></details>')

    pieces.append("</div>")
    return "".join(pieces)


def generate_html(details, all_rows, jtl_name):
    summary = summarize(details, all_rows)
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    passed_rows = summary["total_rows"] - summary["raw_failed"]
    transaction_rate = rate_text(summary["transaction_success"], summary["transaction_total"])

    parts = [
        "<!DOCTYPE html>",
        '<html lang="zh-CN">',
        "<head>",
        '<meta charset="UTF-8">',
        '<meta name="viewport" content="width=device-width, initial-scale=1.0">',
        f"<title>{h(jtl_name)} - 中间件接口失败报告</title>",
        css_block(),
        "</head>",
        "<body><div class=\"wrap\">",
        '<div class="top">',
        "<h1>CRM2.0 中间件接口对比失败报告</h1>",
        f'<div class="sub">数据来源: {h(jtl_name)} | 生成时间: {h(generated_at)}</div>',
        "</div>",
        '<div class="stats primary">',
        f'<div class="stat"><div class="num">{summary["transaction_total"]}</div><div class="name">事务总数</div></div>',
        f'<div class="stat"><div class="num red">{summary["transaction_failed"]}</div><div class="name">失败事务</div></div>',
        f'<div class="stat"><div class="num blue">{transaction_rate}</div><div class="name">事务成功率 = {summary["transaction_success"]}/{summary["transaction_total"]}</div></div>',
        f'<div class="stat"><div class="num red">{summary["fail_field_unique_count"]}</div><div class="name">对比失败字段总数(去重)</div></div>',
        "</div>",
        '<div class="stats secondary">',
        f'<div class="stat"><div class="num">{summary["total_rows"]}</div><div class="name">JTL 总样本</div></div>',
        f'<div class="stat"><div class="num green">{passed_rows}</div><div class="name">JTL 通过样本</div></div>',
        f'<div class="stat"><div class="num red">{summary["raw_failed"]}</div><div class="name">JTL 失败样本</div></div>',
        "</div>",
    ]

    if not details:
        parts.append('<div class="section"><div class="empty">未发现失败样本，当前 JTL 没有需要输出的报错部分。</div></div>')
    else:
        parts.append(render_overview(summary))
        for api, items in iter_api_groups(summary):
            fail_fields = sum(len(d["failed_fields"]) for d in items)
            parts.append(
                '<details class="section api-group">'
                '<summary class="section-title">'
                f'<span class="api-summary-title">接口: <strong>{h(api)}</strong></span>'
                '<span class="api-summary-meta">'
                f'<span class="pill fail">{len(items)} 条失败 / 失败字段明细 {fail_fields} 个</span>'
                '<span class="pill">点击展开</span>'
                '</span>'
                '</summary>'
                '<div class="api-group-body">'
            )
            parts.append(render_api_failed_field_summary(items))
            for i, detail in enumerate(items, 1):
                parts.append(render_sample(detail, i))
            parts.append("</div></details>")

    parts.append("</div></body></html>")
    return "\n".join(parts)


def shorten_text(value, limit=600):
    text = "" if value is None else str(value)
    return text if len(text) <= limit else text[: max(limit - 3, 0)] + "..."


def is_http_url(value):
    try:
        parsed = urlparse(value or "")
        return parsed.scheme in ("http", "https") and bool(parsed.netloc)
    except Exception:
        return False


def is_report_open_url(value):
    try:
        parsed = urlparse(value or "")
        if parsed.scheme in ("http", "https") and parsed.netloc:
            return True
        if parsed.scheme == "file" and parsed.path:
            return True
    except Exception:
        pass
    return False


def local_file_url(path):
    try:
        return Path(os.path.abspath(path)).as_uri()
    except Exception:
        normalized = os.path.abspath(path).replace("\\", "/")
        return "file:///" + quote(normalized, safe="/:")


def is_teams_chat_page_link(value):
    try:
        parsed = urlparse(value or "")
        return "teams.microsoft.com" in parsed.netloc.lower() and "/l/chat/" in parsed.path.lower()
    except Exception:
        return False


def load_config(path):
    if not path:
        return {}
    config_path = os.path.abspath(path)
    if not os.path.isfile(config_path):
        return {}
    try:
        with open(config_path, "r", encoding="utf-8-sig") as f:
            data = json.load(f)
        if isinstance(data, dict):
            return data
    except Exception as exc:
        print(f"配置文件读取失败，已忽略: {config_path} ({exc})")
    return {}


def first_config_value(*values):
    for value in values:
        if value is None:
            continue
        text = str(value).strip()
        if text:
            return text
    return ""


def config_bool(value, default=False):
    if value is None or value == "":
        return default
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in ("1", "true", "yes", "y", "on")


def config_int(value, default):
    try:
        if value is None or value == "":
            return default
        return int(value)
    except Exception:
        return default


def extract_chat_id_from_teams_url(chat_url):
    try:
        parsed = urlparse(chat_url or "")
        parts = [p for p in parsed.path.split("/") if p]
        for index, part in enumerate(parts):
            if part == "chat" and index + 1 < len(parts):
                return unquote(parts[index + 1])
    except Exception:
        pass
    return ""


def resolve_teams_options(args, config):
    chat_url = first_config_value(
        args.teams_chat_url,
        os.environ.get("TEAMS_CHAT_URL"),
        os.environ.get("TM_CHAT_URL"),
        config.get("teams_chat_url"),
        config.get("tm_chat_url"),
        DEFAULT_TEAMS_CHAT_URL,
    )
    return {
        "chat_url": chat_url,
        "chat_id": first_config_value(
            args.teams_chat_id,
            os.environ.get("TEAMS_CHAT_ID"),
            os.environ.get("TM_CHAT_ID"),
            config.get("teams_chat_id"),
            config.get("tm_chat_id"),
            extract_chat_id_from_teams_url(chat_url),
        ),
        "webhook_url": first_config_value(
            args.teams_webhook,
            os.environ.get("TEAMS_WEBHOOK_URL"),
            os.environ.get("TM_WEBHOOK_URL"),
            config.get("teams_webhook_url"),
            config.get("tm_webhook_url"),
        ),
        "report_url": first_config_value(
            args.teams_report_url,
            os.environ.get("TEAMS_REPORT_URL"),
            os.environ.get("TM_REPORT_URL"),
            config.get("teams_report_url"),
            config.get("tm_report_url"),
        ),
        "graph_token": first_config_value(
            args.teams_graph_token,
            os.environ.get("TEAMS_GRAPH_TOKEN"),
            os.environ.get("TM_GRAPH_TOKEN"),
            os.environ.get("GRAPH_ACCESS_TOKEN"),
            config.get("teams_graph_token"),
            config.get("tm_graph_token"),
        ),
        "tenant_id": first_config_value(
            args.teams_tenant_id,
            os.environ.get("TEAMS_TENANT_ID"),
            os.environ.get("TM_TENANT_ID"),
            os.environ.get("AZURE_TENANT_ID"),
            config.get("teams_tenant_id"),
            config.get("tm_tenant_id"),
        ),
        "client_id": first_config_value(
            args.teams_client_id,
            os.environ.get("TEAMS_CLIENT_ID"),
            os.environ.get("TM_CLIENT_ID"),
            os.environ.get("AZURE_CLIENT_ID"),
            config.get("teams_client_id"),
            config.get("tm_client_id"),
        ),
        "scopes": first_config_value(
            args.teams_scopes,
            os.environ.get("TEAMS_GRAPH_SCOPES"),
            os.environ.get("TM_GRAPH_SCOPES"),
            config.get("teams_graph_scopes"),
            config.get("tm_graph_scopes"),
            GRAPH_DEFAULT_SCOPES,
        ),
        "auto_login": (
            args.teams_auto_login
            or config_bool(os.environ.get("TEAMS_AUTO_LOGIN"), False)
            or config_bool(os.environ.get("TM_AUTO_LOGIN"), False)
            or config_bool(config.get("teams_auto_login"), False)
            or config_bool(config.get("tm_auto_login"), False)
        ),
        "token_cache": first_config_value(
            args.teams_token_cache,
            os.environ.get("TEAMS_TOKEN_CACHE"),
            os.environ.get("TM_TOKEN_CACHE"),
            config.get("teams_token_cache"),
            config.get("tm_token_cache"),
            DEFAULT_TOKEN_CACHE_FILE,
        ),
        "upload_folder": first_config_value(
            args.teams_upload_folder,
            os.environ.get("TEAMS_UPLOAD_FOLDER"),
            os.environ.get("TM_UPLOAD_FOLDER"),
            config.get("teams_upload_folder"),
            config.get("tm_upload_folder"),
            "JMeterReports",
        ),
        "send_file": (
            args.teams_send_file
            or config_bool(os.environ.get("TEAMS_SEND_FILE"), False)
            or config_bool(os.environ.get("TM_SEND_FILE"), False)
            or config_bool(config.get("teams_send_file"), False)
            or config_bool(config.get("tm_send_file"), False)
        ),
        "timeout": config_int(args.teams_timeout if args.teams_timeout is not None else config.get("teams_timeout"), 20),
        "strict": args.teams_strict or config_bool(config.get("teams_strict"), False),
    }


def build_teams_payload(details, all_rows, jtl_name, output_path, report_url=None, chat_url=None):
    summary = summarize(details, all_rows)
    passed_rows = summary["total_rows"] - summary["raw_failed"]
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    transaction_rate = rate_text(summary["transaction_success"], summary["transaction_total"])
    status_text = "发现失败，请查看报告" if details else "未发现失败样本"
    status_color = "Attention" if details else "Good"
    report_open_url = report_url or local_file_url(output_path)
    report_link_text = "HTML报告: [点击打开报告]({})".format(report_open_url)

    top_apis = []
    sorted_api_items = list(iter_api_groups(summary))
    for api, items in sorted_api_items[:5]:
        api_fail_count = sum(len(d["failed_fields"]) for d in items)
        field_counts = {}
        for item in summary["fail_fields"]:
            if item["api_name"] == api:
                field_counts[item["field"]] = field_counts.get(item["field"], 0) + 1
        top_fields = sorted(field_counts.items(), key=lambda x: (-x[1], x[0]))[:3]
        top_field_text = "、".join(f"{field}({count})" for field, count in top_fields) if top_fields else "无字段失败明细"
        top_apis.append(f"- {api}: {len(items)}条失败 / 失败字段明细{api_fail_count}个 / {top_field_text}")
    remaining_api_count = max(len(sorted_api_items) - len(top_apis), 0)
    if remaining_api_count:
        top_apis.append(f"- 其余 {remaining_api_count} 个接口请打开 HTML 报告查看")

    facts = [
        {"title": "JTL文件", "value": shorten_text(jtl_name, 180)},
        {"title": "事务总数", "value": str(summary["transaction_total"])},
        {"title": "失败事务", "value": str(summary["transaction_failed"])},
        {"title": "事务成功率", "value": f"{transaction_rate} ({summary['transaction_success']}/{summary['transaction_total']})"},
        {"title": "对比失败字段(去重)", "value": str(summary["fail_field_unique_count"])},
        {"title": "JTL总样本", "value": str(summary["total_rows"])},
        {"title": "JTL通过样本", "value": str(passed_rows)},
        {"title": "JTL失败样本", "value": str(summary["raw_failed"])},
    ]

    body = [
        {
            "type": "TextBlock",
            "text": "CRM2.0 中间件接口对比失败报告",
            "weight": "Bolder",
            "size": "Large",
            "wrap": True,
        },
        {
            "type": "TextBlock",
            "text": status_text,
            "weight": "Bolder",
            "color": status_color,
            "wrap": True,
        },
        {
            "type": "TextBlock",
            "text": report_link_text,
            "weight": "Bolder",
            "wrap": True,
            "spacing": "Small",
        },
        {
            "type": "TextBlock",
            "text": f"生成时间: {generated_at}",
            "isSubtle": True,
            "spacing": "Small",
            "wrap": True,
        },
        {"type": "FactSet", "facts": facts},
    ]

    if top_apis:
        body.append({
            "type": "TextBlock",
            "text": "**接口失败统计 / 高频字段**\n" + shorten_text("\n".join(top_apis), 1200),
            "wrap": True,
            "spacing": "Medium",
        })

    card = {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.2",
        "body": body,
    }
    if is_report_open_url(report_open_url):
        card.setdefault("actions", []).append({
            "type": "Action.OpenUrl",
            "title": "打开HTML报告",
            "url": report_open_url,
        })
    if chat_url and is_http_url(chat_url):
        card.setdefault("actions", []).append({
            "type": "Action.OpenUrl",
            "title": "打开Teams会话",
            "url": chat_url,
        })

    return {
        "type": "message",
        "attachments": [{
            "contentType": "application/vnd.microsoft.card.adaptive",
            "contentUrl": None,
            "content": card,
        }],
    }


def send_teams_webhook(webhook_url, payload, timeout=20):
    if is_teams_chat_page_link(webhook_url):
        raise ValueError(
            "当前填写的是 Teams 聊天页面链接，不是 Webhook 接收地址。"
            "请在 Teams Workflows/Incoming Webhook 中创建 HTTP POST 地址后再填写。"
        )
    if not is_http_url(webhook_url):
        raise ValueError("Teams Webhook 地址必须是 http/https URL。")

    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urlrequest.Request(
        webhook_url,
        data=data,
        headers={
            "Content-Type": "application/json; charset=utf-8",
            "User-Agent": f"jtl-fail-report/{SCRIPT_VERSION}",
        },
        method="POST",
    )
    try:
        with urlrequest.urlopen(req, timeout=timeout) as resp:
            body = resp.read().decode("utf-8", errors="replace")
            return resp.status, body
    except urlerror.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTP {exc.code}: {shorten_text(body.strip(), 1000)}") from exc
    except urlerror.URLError as exc:
        raise RuntimeError(f"网络请求失败: {exc.reason}") from exc


class OAuthError(RuntimeError):
    def __init__(self, code, description, response=None):
        self.code = code or "oauth_error"
        self.description = description or ""
        self.response = response or {}
        super().__init__(f"{self.code}: {self.description}".strip(": "))


def oauth_form_post(url, form, timeout=30):
    data = urlencode(form).encode("utf-8")
    req = urlrequest.Request(
        url,
        data=data,
        headers={
            "Content-Type": "application/x-www-form-urlencoded",
            "User-Agent": f"jtl-fail-report/{SCRIPT_VERSION}",
        },
        method="POST",
    )
    try:
        with urlrequest.urlopen(req, timeout=timeout) as resp:
            raw = resp.read().decode("utf-8", errors="replace")
            return json.loads(raw) if raw else {}
    except urlerror.HTTPError as exc:
        raw = exc.read().decode("utf-8", errors="replace")
        try:
            data = json.loads(raw)
            raise OAuthError(data.get("error"), data.get("error_description") or raw, data) from exc
        except json.JSONDecodeError:
            raise RuntimeError(f"OAuth HTTP {exc.code}: {shorten_text(raw.strip(), 1200)}") from exc
    except urlerror.URLError as exc:
        raise RuntimeError(f"OAuth 网络请求失败: {exc.reason}") from exc


def token_cache_file(path):
    if not path:
        return ""
    return os.path.abspath(path)


def normalize_scopes(scopes):
    return " ".join((scopes or GRAPH_DEFAULT_SCOPES).split())


def load_token_cache(path):
    path = token_cache_file(path)
    if not path or not os.path.isfile(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_token_cache(path, token, tenant_id, client_id, scopes):
    path = token_cache_file(path)
    if not path:
        return
    expires_in = int(token.get("expires_in") or 3600)
    data = {
        "tenant_id": tenant_id,
        "client_id": client_id,
        "scopes": normalize_scopes(scopes),
        "access_token": token.get("access_token", ""),
        "refresh_token": token.get("refresh_token", ""),
        "expires_at": int(time.time()) + expires_in,
        "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def cached_graph_token(path, tenant_id, client_id, scopes):
    data = load_token_cache(path)
    if not data:
        return ""
    if data.get("tenant_id") != tenant_id or data.get("client_id") != client_id:
        return ""
    if normalize_scopes(data.get("scopes")) != normalize_scopes(scopes):
        return ""
    if int(data.get("expires_at") or 0) <= int(time.time()) + 300:
        return ""
    return data.get("access_token", "")


def refresh_graph_token(cache_path, tenant_id, client_id, scopes, timeout=30):
    data = load_token_cache(cache_path)
    refresh_token = data.get("refresh_token")
    if not refresh_token:
        return ""
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    try:
        token = oauth_form_post(
            token_url,
            {
                "client_id": client_id,
                "grant_type": "refresh_token",
                "refresh_token": refresh_token,
                "scope": normalize_scopes(scopes),
            },
            timeout=timeout,
        )
    except Exception:
        return ""
    if token.get("access_token"):
        if not token.get("refresh_token"):
            token["refresh_token"] = refresh_token
        save_token_cache(cache_path, token, tenant_id, client_id, scopes)
        return token["access_token"]
    return ""


def login_graph_device_code(tenant_id, client_id, scopes, cache_path, timeout=30):
    scope_text = normalize_scopes(scopes)
    device_code_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/devicecode"
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

    device = oauth_form_post(
        device_code_url,
        {"client_id": client_id, "scope": scope_text},
        timeout=timeout,
    )
    print("\nMicrosoft Graph 登录:")
    print("  请按下面提示完成登录授权。")
    if device.get("message"):
        print(f"  {device['message']}")
    else:
        print(f"  打开: {device.get('verification_uri') or 'https://microsoft.com/devicelogin'}")
        print(f"  输入验证码: {device.get('user_code')}")

    interval = int(device.get("interval") or 5)
    expires_at = time.time() + int(device.get("expires_in") or 900)
    while time.time() < expires_at:
        time.sleep(interval)
        try:
            token = oauth_form_post(
                token_url,
                {
                    "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
                    "client_id": client_id,
                    "device_code": device["device_code"],
                },
                timeout=timeout,
            )
            if token.get("access_token"):
                save_token_cache(cache_path, token, tenant_id, client_id, scopes)
                print("  Graph 登录成功，已获取 access token。")
                return token["access_token"]
        except OAuthError as exc:
            if exc.code == "authorization_pending":
                continue
            if exc.code == "slow_down":
                interval += 5
                continue
            raise
    raise RuntimeError("Graph Device Code 登录超时，请重新运行脚本。")


def resolve_graph_access_token(teams_options):
    if teams_options["graph_token"]:
        return teams_options["graph_token"]

    if not teams_options["auto_login"]:
        raise ValueError(
            "已要求发送 HTML 文件到 Teams，但未配置 Graph token。"
            "请配置 teams_graph_token，或环境变量 TEAMS_GRAPH_TOKEN / GRAPH_ACCESS_TOKEN；"
            "也可以配置 teams_tenant_id + teams_client_id 后启用 teams_auto_login。"
        )

    tenant_id = teams_options["tenant_id"]
    client_id = teams_options["client_id"]
    if not tenant_id or not client_id:
        raise ValueError(
            "已启用 teams_auto_login，但缺少 teams_tenant_id 或 teams_client_id。"
            "请先在 Microsoft Entra ID 注册应用，并把 Directory tenant ID 和 Application client ID 填入配置。"
        )

    cached = cached_graph_token(teams_options["token_cache"], tenant_id, client_id, teams_options["scopes"])
    if cached:
        print("\nMicrosoft Graph 登录: 使用本地 token 缓存")
        return cached

    refreshed = refresh_graph_token(
        teams_options["token_cache"],
        tenant_id,
        client_id,
        teams_options["scopes"],
        timeout=teams_options["timeout"],
    )
    if refreshed:
        print("\nMicrosoft Graph 登录: 已通过 refresh token 自动刷新")
        return refreshed

    return login_graph_device_code(
        tenant_id,
        client_id,
        teams_options["scopes"],
        teams_options["token_cache"],
        timeout=teams_options["timeout"],
    )


def graph_request(method, url, access_token, timeout=30, json_body=None, data=None, content_type="application/json"):
    headers = {
        "Authorization": f"Bearer {access_token}",
        "User-Agent": f"jtl-fail-report/{SCRIPT_VERSION}",
    }
    if json_body is not None:
        data = json.dumps(json_body, ensure_ascii=False).encode("utf-8")
        headers["Content-Type"] = "application/json; charset=utf-8"
    elif content_type:
        headers["Content-Type"] = content_type

    req = urlrequest.Request(url, data=data, headers=headers, method=method)
    try:
        with urlrequest.urlopen(req, timeout=timeout) as resp:
            raw = resp.read().decode("utf-8", errors="replace")
            if not raw:
                return {}
            return json.loads(raw)
    except urlerror.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Graph HTTP {exc.code}: {shorten_text(body.strip(), 1200)}") from exc
    except urlerror.URLError as exc:
        raise RuntimeError(f"Graph 网络请求失败: {exc.reason}") from exc


def ensure_graph_folder(access_token, folder_path, timeout=30):
    folder_path = (folder_path or "").strip().strip("/\\")
    if not folder_path:
        return "root"

    parent_id = "root"
    current_parts = []
    for raw_part in re.split(r"[\\/]+", folder_path):
        part = raw_part.strip()
        if not part:
            continue
        current_parts.append(part)
        graph_path = quote("/".join(current_parts), safe="/")
        get_url = f"{GRAPH_BASE_URL}/me/drive/root:/{graph_path}"
        try:
            item = graph_request("GET", get_url, access_token, timeout=timeout, content_type="")
            parent_id = item.get("id") or parent_id
            continue
        except RuntimeError as exc:
            if "Graph HTTP 404" not in str(exc):
                raise

        if parent_id == "root":
            create_url = f"{GRAPH_BASE_URL}/me/drive/root/children"
        else:
            create_url = f"{GRAPH_BASE_URL}/me/drive/items/{quote(parent_id, safe='')}/children"
        item = graph_request(
            "POST",
            create_url,
            access_token,
            timeout=timeout,
            json_body={
                "name": part,
                "folder": {},
                "@microsoft.graph.conflictBehavior": "fail",
            },
        )
        parent_id = item.get("id") or parent_id
    return parent_id


def upload_report_to_onedrive(access_token, file_path, folder_path, timeout=60):
    ensure_graph_folder(access_token, folder_path, timeout=timeout)
    file_name = os.path.basename(file_path)
    remote_path = "/".join(
        part.strip("/\\")
        for part in (folder_path, file_name)
        if part and part.strip("/\\")
    )
    if not remote_path:
        remote_path = file_name
    upload_url = f"{GRAPH_BASE_URL}/me/drive/root:/{quote(remote_path, safe='/')}:/content"
    with open(file_path, "rb") as f:
        data = f.read()
    content_type = "text/html; charset=utf-8" if file_name.lower().endswith((".html", ".htm")) else "application/octet-stream"
    return graph_request("PUT", upload_url, access_token, timeout=timeout, data=data, content_type=content_type)


def create_onedrive_share_link(access_token, item_id, timeout=30):
    url = f"{GRAPH_BASE_URL}/me/drive/items/{quote(item_id, safe='')}/createLink"
    result = graph_request(
        "POST",
        url,
        access_token,
        timeout=timeout,
        json_body={"type": "view", "scope": "organization"},
    )
    return ((result.get("link") or {}).get("webUrl") or "").strip()


def get_drive_item_attachment_info(access_token, item_id, timeout=30):
    url = (
        f"{GRAPH_BASE_URL}/me/drive/items/{quote(item_id, safe='')}"
        "?$select=id,name,webUrl,webDavUrl,eTag"
    )
    return graph_request("GET", url, access_token, timeout=timeout, content_type="")


def attachment_id_from_etag(etag):
    match = re.search(r"[{(]?([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})[})]?", etag or "")
    return match.group(1) if match else str(uuid.uuid4())


def post_teams_file_message(access_token, chat_id, file_name, content_url, attachment_id=None, timeout=30):
    attachment_id = attachment_id or str(uuid.uuid4())
    chat_url = f"{GRAPH_BASE_URL}/chats/{quote(chat_id, safe='')}/messages"
    payload = {
        "body": {
            "contentType": "html",
            "content": (
                "中间件接口对比失败报告已生成，请查看附件："
                f'<attachment id="{attachment_id}"></attachment>'
            ),
        },
        "attachments": [
            {
                "id": attachment_id,
                "contentType": "reference",
                "contentUrl": content_url,
                "name": file_name,
            }
        ],
    }
    return graph_request("POST", chat_url, access_token, timeout=timeout, json_body=payload)


def push_teams_report(teams_options, details, rows, jtl_name, output_path):
    webhook_url = teams_options["webhook_url"]
    if not webhook_url:
        return None

    report_url = teams_options["report_url"]
    if report_url and not is_report_open_url(report_url):
        raise ValueError("--teams-report-url / TEAMS_REPORT_URL 必须是 http/https 或 file:/// 可打开链接。")

    payload = build_teams_payload(
        details,
        rows,
        jtl_name,
        output_path,
        report_url=report_url,
        chat_url=teams_options["chat_url"],
    )
    return send_teams_webhook(webhook_url, payload, timeout=teams_options["timeout"])


def push_teams_file_report(teams_options, output_path):
    if not teams_options["send_file"]:
        return None
    access_token = resolve_graph_access_token(teams_options)
    chat_id = teams_options["chat_id"]
    if not chat_id:
        raise ValueError("未能从 Teams 链接解析 chat id，请配置 teams_chat_id。")
    if not os.path.isfile(output_path):
        raise ValueError(f"HTML 报告文件不存在: {output_path}")

    uploaded = upload_report_to_onedrive(
        access_token,
        output_path,
        teams_options["upload_folder"],
        timeout=max(teams_options["timeout"], 60),
    )
    item_id = uploaded.get("id")
    if not item_id:
        raise RuntimeError(f"Graph 上传成功但未返回 driveItem id: {uploaded}")

    item_info = get_drive_item_attachment_info(access_token, item_id, timeout=teams_options["timeout"])
    content_url = item_info.get("webDavUrl") or item_info.get("webUrl")
    share_url = ""
    try:
        share_url = create_onedrive_share_link(access_token, item_id, timeout=teams_options["timeout"])
    except Exception:
        share_url = ""
    content_url = content_url or share_url
    if not content_url:
        raise RuntimeError(f"Graph 上传成功但未返回可用于 Teams 附件的文件链接: {item_info}")

    message = post_teams_file_message(
        access_token,
        chat_id,
        item_info.get("name") or os.path.basename(output_path),
        content_url,
        attachment_id=attachment_id_from_etag(item_info.get("eTag")),
        timeout=teams_options["timeout"],
    )
    return {
        "share_url": share_url,
        "content_url": content_url,
        "message_id": message.get("id", ""),
        "chat_id": chat_id,
        "uploaded_name": item_info.get("name") or uploaded.get("name", os.path.basename(output_path)),
    }


def teams_webhook_failure_hint(exc):
    text = str(exc)
    if "DirectApiAuthorizationRequired" in text:
        return (
            "提示: 当前 Power Automate URL 需要 OAuth 授权，不是可直接 POST 的 Workflows Webhook URL。"
            "请在 Teams Workflows 中使用 webhook 接收模板，并复制生成的匿名 HTTP POST URL；"
            "或调整触发器权限为允许外部请求。"
        )
    if "teams.microsoft.com/l/chat/" in text:
        return "提示: Teams 聊天页面链接不能直接作为 Webhook；需要复制 Workflows/Incoming Webhook 生成的 HTTP POST URL。"
    return "提示: 请确认 teams_webhook_url 是 Teams Workflows 生成的 HTTP POST URL，且该 Workflow 已启用。"


def parse_args():
    parser = argparse.ArgumentParser(
        description="根据 JMeter .jtl 生成只包含失败/报错详情的 HTML 报告。",
        formatter_class=argparse.RawTextHelpFormatter,
    )
    parser.add_argument("jtl", nargs="?", help="输入 JTL 文件路径，省略时自动选择脚本目录下最新 .jtl")
    parser.add_argument("-o", "--output", help="输出 HTML 文件路径，默认: <JTL文件名>_失败报告.html")
    parser.add_argument("--no-open", action="store_true", help="生成后不自动打开浏览器")
    parser.add_argument("--config", default=DEFAULT_CONFIG_FILE, help=f"配置文件路径，默认: {DEFAULT_CONFIG_FILE}")
    parser.add_argument(
        "--teams-chat-url",
        help="TM/Teams 聊天页面链接；也可用配置项 teams_chat_url 或环境变量 TEAMS_CHAT_URL/TM_CHAT_URL。",
    )
    parser.add_argument(
        "--teams-chat-id",
        help="Microsoft Graph chat id；默认尝试从 teams_chat_url 中解析，也可用配置项 teams_chat_id 或环境变量 TEAMS_CHAT_ID/TM_CHAT_ID。",
    )
    parser.add_argument(
        "--teams-webhook",
        help="Teams Workflows/Incoming Webhook 的 HTTP POST 地址；也可用配置项 teams_webhook_url 或环境变量 TEAMS_WEBHOOK_URL/TM_WEBHOOK_URL。",
    )
    parser.add_argument(
        "--teams-report-url",
        help="HTML报告链接，支持 http/https 或 file:/// 本地文件链接；也可用配置项 teams_report_url 或环境变量 TEAMS_REPORT_URL/TM_REPORT_URL。",
    )
    parser.add_argument(
        "--teams-send-file",
        action="store_true",
        help="使用 Microsoft Graph 上传当前 HTML 报告到 OneDrive 并作为文件附件发送到 Teams chat。",
    )
    parser.add_argument(
        "--teams-graph-token",
        help="Microsoft Graph access token；也可用配置项 teams_graph_token 或环境变量 TEAMS_GRAPH_TOKEN/GRAPH_ACCESS_TOKEN。",
    )
    parser.add_argument(
        "--teams-auto-login",
        action="store_true",
        help="未配置 Graph token 时，使用 Microsoft Device Code Flow 登录并自动获取 token。",
    )
    parser.add_argument(
        "--teams-tenant-id",
        help="Microsoft Entra Directory tenant ID；也可用配置项 teams_tenant_id 或环境变量 TEAMS_TENANT_ID/AZURE_TENANT_ID。",
    )
    parser.add_argument(
        "--teams-client-id",
        help="Microsoft Entra Application client ID；也可用配置项 teams_client_id 或环境变量 TEAMS_CLIENT_ID/AZURE_CLIENT_ID。",
    )
    parser.add_argument(
        "--teams-scopes",
        help=f"Graph OAuth scopes，默认: {GRAPH_DEFAULT_SCOPES}",
    )
    parser.add_argument(
        "--teams-token-cache",
        help=f"Graph token 缓存文件，默认: {DEFAULT_TOKEN_CACHE_FILE}",
    )
    parser.add_argument(
        "--teams-upload-folder",
        help="Graph 上传到当前用户 OneDrive 的目录，默认 JMeterReports；也可用配置项 teams_upload_folder。",
    )
    parser.add_argument("--teams-timeout", type=int, help="Teams 推送超时时间，默认 20 秒。")
    parser.add_argument("--teams-strict", action="store_true", help="Teams 推送失败时返回非 0 退出码。")
    return parser.parse_args()


def main():
    args = parse_args()
    config = load_config(args.config)
    teams_options = resolve_teams_options(args, config)
    print("=" * 72)
    print(f"  CRM2.0 中间件接口 JTL 失败报告生成器 v{SCRIPT_VERSION}")
    print("=" * 72)
    if teams_options["chat_url"]:
        print(f"\nTM/Teams 会话入口: {teams_options['chat_url']}")
    if teams_options["chat_id"]:
        print(f"Teams chat id: {teams_options['chat_id']}")
    if teams_options["webhook_url"]:
        print("Teams Webhook: 已配置，报告生成后将推送摘要卡片")
    else:
        print("Teams Webhook: 未配置，仅生成本地 HTML 报告")
    if teams_options["send_file"]:
        print("Teams 文件发送: 已启用，报告生成后将通过 Microsoft Graph 发送 HTML 文件")

    jtl_path = find_jtl_file(args.jtl)
    jtl_name = os.path.basename(jtl_path)
    jtl_dir = os.path.dirname(jtl_path)
    output_path = args.output
    if not output_path:
        output_path = os.path.join(jtl_dir, f"{os.path.splitext(jtl_name)[0]}_失败报告.html")
    else:
        output_path = os.path.abspath(output_path)

    print(f"\nJTL 文件: {jtl_path}")
    print(f"文件大小: {os.path.getsize(jtl_path) / 1024:.1f} KB")

    rows = parse_jtl(jtl_path)
    details = extract_failures(rows)
    summary = summarize(details, rows)
    passed_rows = summary["total_rows"] - summary["raw_failed"]
    transaction_rate = rate_text(summary["transaction_success"], summary["transaction_total"])

    print("\n解析结果:")
    print(f"  事务总数: {summary['transaction_total']}")
    print(f"  失败事务: {summary['transaction_failed']}")
    print(f"  事务成功率: {transaction_rate} ({summary['transaction_success']}/{summary['transaction_total']})")
    print(f"  对比失败字段总数(去重): {summary['fail_field_unique_count']}")
    print(f"  JTL 总样本: {summary['total_rows']}")
    print(f"  JTL 通过样本: {passed_rows}")
    print(f"  JTL 失败样本: {summary['raw_failed']}")

    html = generate_html(details, rows, jtl_name)
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    with open(output_path, "w", encoding="utf-8", newline="") as f:
        f.write(html)

    print("\n报告已生成:")
    print(f"  {output_path}")
    print(f"  大小: {os.path.getsize(output_path) / 1024:.1f} KB")

    teams_failed = False
    try:
        teams_result = push_teams_report(teams_options, details, rows, jtl_name, output_path)
        if teams_result:
            status, body = teams_result
            print("\nTeams 推送:")
            print(f"  已发送到 Teams Webhook，HTTP状态: {status}")
            if body and body.strip() not in ("1", "ok", "OK"):
                print(f"  Teams返回: {shorten_text(body.strip(), 300)}")
    except Exception as exc:
        teams_failed = True
        print("\nTeams 推送失败:")
        print(f"  {exc}")
        print(f"  {teams_webhook_failure_hint(exc)}")

    try:
        file_result = push_teams_file_report(teams_options, output_path)
        if file_result:
            print("\nTeams 文件发送:")
            print(f"  已上传并发送: {file_result['uploaded_name']}")
            print(f"  chat id: {file_result['chat_id']}")
            if file_result["message_id"]:
                print(f"  message id: {file_result['message_id']}")
            if file_result["share_url"]:
                print(f"  分享链接: {file_result['share_url']}")
            else:
                print(f"  附件引用链接: {file_result['content_url']}")
    except Exception as exc:
        teams_failed = True
        print("\nTeams 文件发送失败:")
        print(f"  {exc}")
        print("  提示: 发送本地 HTML 文件需要 Microsoft Graph token，权限通常需要 Files.ReadWrite 和 ChatMessage.Send。")

    if not args.no_open:
        try:
            import webbrowser
            webbrowser.open(output_path)
            print("  已尝试自动打开浏览器")
        except Exception:
            print("  浏览器自动打开失败，请手动打开 HTML 文件")

    print("=" * 72)
    if teams_failed and teams_options["strict"]:
        sys.exit(2)


if __name__ == "__main__":
    main()
