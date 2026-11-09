# -*- coding: utf-8 -*-
"""
本地模拟火山方舟 chat/completions 接口
====================================
用于在不消耗真实 API 额度的情况下，演练「AI 用例评审脚本」的完整流程
（解析 -> 调用 -> 评分 -> 报告输出）。

用法：
    python mock_ark_server.py --port 8000
    然后：
    python review_llm_testcases.py --input 样例用例.xlsx \
        --base-url http://127.0.0.1:8000 --api-key mock

说明：mock 会根据用例文本里的几个特征（缺预期结果、步骤模糊等）做一个
简单的确定性评分，只用于验证流程，不代表真实模型的评审质量。
"""

import argparse
import json
import re
from http.server import BaseHTTPRequestHandler, HTTPServer

DIMENSIONS = [
    ("需求可追溯性", 0.10),
    ("完整性", 0.15),
    ("准确性", 0.15),
    ("可执行性", 0.10),
    ("预期结果可验证性", 0.15),
    ("边界与异常覆盖", 0.10),
    ("独立性与去重", 0.10),
    ("优先级合理性", 0.05),
    ("命名与可读性", 0.05),
    ("数据与安全考虑", 0.05),
]


_seen = {}
_last_req_time = [0.0]   # 用于跨轮次重置重复指纹


def make_review(case_text):
    """根据用例文本特征生成确定性评审结果（含典型缺陷打分）。"""
    import time as _t
    now = _t.time()
    if now - _last_req_time[0] > 10:   # 间隔超过10秒视为新一轮评审，清空指纹
        _seen.clear()
    _last_req_time[0] = now

    scores = {name: 9 for name, _ in DIMENSIONS}   # 无缺陷基线 9 分
    issues = []

    has_expected = "预期结果：" in case_text
    has_steps = "测试步骤：" in case_text
    has_req = "关联需求：" in case_text
    has_pre = "前置条件：" in case_text

    if not has_expected:
        scores["预期结果可验证性"] = 2
        issues.append({
            "level": "严重", "dimension": "预期结果可验证性",
            "problem": "缺少预期结果，无法判断用例执行是否通过",
            "suggestion": "补充明确的预期结果，给出具体值或可观察状态，例如“提示‘用户名不能为空’且停留登录页”",
        })
        scores["完整性"] = 6
        issues.append({
            "level": "一般", "dimension": "完整性",
            "problem": "用例要素缺失：未提供预期结果",
            "suggestion": "补齐预期结果字段后再提交评审",
        })
    if not has_req:
        scores["需求可追溯性"] = 3
        issues.append({
            "level": "一般", "dimension": "需求可追溯性",
            "problem": "未关联需求编号/需求描述，无法追溯用例覆盖的需求点",
            "suggestion": "补充需求编号（如 REQ-LOGIN-001）或需求描述，确保用例可追溯",
        })
    if not has_pre:
        scores["完整性"] = min(scores["完整性"], 6)
        issues.append({
            "level": "轻微", "dimension": "完整性",
            "problem": "缺少前置条件说明",
            "suggestion": "补充前置条件，例如账号已注册、系统已部署等",
        })
    if re.search(r"(随便|酌情|适当|进行一些|若干|部分)", case_text):
        scores["可执行性"] = 3
        issues.append({
            "level": "一般", "dimension": "可执行性",
            "problem": "步骤中存在含糊表述（如“随便/酌情/适当”），测试人员无法按统一标准执行",
            "suggestion": "将含糊表述改为具体数值或明确动作，例如“输入 11 位手机号”“点击‘登录’按钮”",
        })
    if "锁定" in case_text and "未锁定" not in case_text \
            and "次" not in case_text and "次数" not in case_text:
        scores["边界与异常覆盖"] = 3
        issues.append({
            "level": "一般", "dimension": "边界与异常覆盖",
            "problem": "提到“锁定”但未给出锁定阈值（次数/时长），边界条件不明确",
            "suggestion": "明确锁定阈值，如“连续输错 5 次锁定 15 分钟”，并补充第 4 次与第 5 次的边界用例",
        })

    # 重复用例检测：按「测试步骤」的文本指纹判断
    import hashlib
    step_m = re.search(r"测试步骤：(.+)", case_text)
    finger_src = step_m.group(1) if step_m else ""
    finger = hashlib.md5(re.sub(r"\s", "", finger_src).encode("utf-8")).hexdigest()
    if finger in _seen:
        scores["独立性与去重"] = 3
        issues.append({
            "level": "严重", "dimension": "独立性与去重",
            "problem": "与已评审的用例内容重复，属于冗余用例",
            "suggestion": "删除该用例，或将差异点（不同前置、数据、断言）明确化后再保留",
        })
    _seen[finger] = True

    overall = round(sum(scores[n] * w for n, w in DIMENSIONS), 1)
    verdict = "通过" if overall >= 8.5 else ("需修改" if overall >= 6.5 else "不通过")
    summary = ("整体结构完整、意图清晰，主要问题集中在"
               + "；".join(i["problem"] for i in issues[:2])
               if issues else "用例结构完整，维度表现均衡，建议保持该质量基线")
    return {
        "overall_score": overall,
        "scores": scores,
        "verdict": verdict,
        "summary": summary,
        "issues": issues,
    }


class Handler(BaseHTTPRequestHandler):
    def do_POST(self):
        length = int(self.headers.get("Content-Length", 0))
        body = json.loads(self.rfile.read(length).decode("utf-8"))
        messages = body.get("messages", [])
        user = next((m["content"] for m in messages if m["role"] == "user"), "")
        review = make_review(user)
        resp = {
            "id": "mock-chatcmpl",
            "object": "chat.completion",
            "model": body.get("model", "mock-model"),
            "choices": [{
                "index": 0,
                "message": {"role": "assistant",
                            "content": json.dumps(review, ensure_ascii=False)},
                "finish_reason": "stop",
            }],
            "usage": {"prompt_tokens": 100, "completion_tokens": 100,
                      "total_tokens": 200},
        }
        data = json.dumps(resp, ensure_ascii=False).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def log_message(self, fmt, *args):
        print(f"[mock] {self.address_string()} {fmt % args}", flush=True)


def main():
    p = argparse.ArgumentParser(description="本地 mock 方舟 /chat/completions")
    p.add_argument("--port", type=int, default=8000)
    args = p.parse_args()
    print(f"mock 接口已启动: http://127.0.0.1:{args.port}/chat/completions", flush=True)
    print("用 Ctrl+C 停止。", flush=True)
    HTTPServer(("127.0.0.1", args.port), Handler).serve_forever()


if __name__ == "__main__":
    main()
