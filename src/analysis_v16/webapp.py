from __future__ import annotations

import argparse
import json
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

from .reporter import build_workbench_pages, write_jsonl, write_markdown
from .review_learning import (
    apply_review_decision,
    build_learning_outputs,
    build_review_batch,
    load_review_decisions,
    validate_review_payload,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="v1.6 AI 一线情报工作台本地服务")
    parser.add_argument("--data", required=True, help="data/output/insights/v1.6 目录")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8766)
    return parser.parse_args()


def serve(data_dir: Path, host: str, port: int) -> None:
    server = ThreadingHTTPServer((host, port), _handler(data_dir))
    print(f"http://{host}:{port}/overview")
    server.serve_forever()


def _handler(data_dir: Path):
    class Handler(BaseHTTPRequestHandler):
        def do_GET(self) -> None:
            path = urlparse(self.path).path
            page_map = {
                "/": "overview.html",
                "/overview": "overview.html",
                "/trends": "trends.html",
                "/sales": "sales.html",
                "/insights": "insights.html",
                "/review": "review.html",
                "/evidence": "evidence.html",
            }
            page_name = page_map.get(path, path.strip("/"))
            if page_name not in {"overview.html", "trends.html", "sales.html", "insights.html", "review.html", "evidence.html"}:
                self.send_error(HTTPStatus.NOT_FOUND)
                return
            content = _build_pages(data_dir).get(page_name)
            if content is None:
                self.send_error(HTTPStatus.NOT_FOUND)
                return
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(content.encode("utf-8"))

        def do_POST(self) -> None:
            if urlparse(self.path).path != "/api/v16-review-decisions":
                self.send_error(HTTPStatus.NOT_FOUND)
                return
            length = int(self.headers.get("Content-Length", "0"))
            payload = json.loads(self.rfile.read(length) or b"{}")
            review_dir = data_dir / "review"
            batch = _read_jsonl(review_dir / "review_batch.jsonl")
            task_map = {str(row.get("task_id", "")): row for row in batch}
            task = task_map.get(str(payload.get("task_id", "")))
            if not task:
                self.send_error(HTTPStatus.BAD_REQUEST, "invalid task_id")
                return
            reviewed_fields = dict(payload.get("reviewed_fields", {}))
            errors = validate_review_payload(task, reviewed_fields)
            if errors:
                self.send_response(HTTPStatus.BAD_REQUEST)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.end_headers()
                self.wfile.write(json.dumps({"ok": False, "error": "；".join(errors)}, ensure_ascii=False).encode("utf-8"))
                return
            decision = apply_review_decision(
                review_dir=review_dir,
                task=task,
                reviewed_fields=reviewed_fields,
                reviewer=str(payload.get("reviewer", "")).strip() or "wales",
                review_comment=str(payload.get("review_comment", "")).strip(),
            )
            _refresh_learning_outputs(data_dir)
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.end_headers()
            self.wfile.write(json.dumps({"ok": True, "decision": decision}, ensure_ascii=False).encode("utf-8"))

        def log_message(self, format: str, *args) -> None:  # noqa: A003
            return

    return Handler


def _refresh_learning_outputs(data_dir: Path) -> None:
    review_dir = data_dir / "review"
    facts = _read_jsonl(data_dir / "normalized" / "business_question_facts.jsonl")
    decisions = load_review_decisions(review_dir / "review_decisions.jsonl")
    batch = build_review_batch(facts, decisions)
    learning_summary, rule_candidates, prompt_candidates, label_candidates, golden_set = build_learning_outputs(decisions, batch)
    write_jsonl(review_dir / "review_batch.jsonl", batch)
    write_markdown(review_dir / "learning_summary.md", learning_summary)
    write_jsonl(review_dir / "rule_candidates.jsonl", rule_candidates)
    write_jsonl(review_dir / "prompt_candidates.jsonl", prompt_candidates)
    write_jsonl(review_dir / "label_gap_candidates.jsonl", label_candidates)
    write_jsonl(review_dir / "golden_set.jsonl", golden_set)


def _build_pages(data_dir: Path):
    normalized = data_dir / "normalized"
    summary = _read_json(normalized / "business_question_summary.json", {})
    facts = _read_jsonl(normalized / "business_question_facts.jsonl")
    insights = _read_jsonl(normalized / "business_insights.jsonl")
    dashboard = _read_json(normalized / "dashboard_snapshot.json", {})
    trend_cube = _read_json(normalized / "trend_cube.json", [])
    profiles = _read_jsonl(normalized / "salesperson_profile.jsonl")
    review_batch = _read_jsonl(data_dir / "review" / "review_batch.jsonl")
    return build_workbench_pages(
        dashboard_snapshot=dashboard,
        business_summary=summary,
        business_facts=facts,
        business_insights=insights,
        review_batch=review_batch,
        trend_cube=trend_cube,
        salesperson_profiles=profiles,
    )


def _read_json(path: Path, default):
    if not path.exists():
        return default
    return json.loads(path.read_text(encoding="utf-8"))


def _read_jsonl(path: Path):
    if not path.exists():
        return []
    return [json.loads(line) for line in path.read_text(encoding="utf-8").splitlines() if line.strip()]


def main() -> None:
    args = parse_args()
    serve(Path(args.data), args.host, args.port)


if __name__ == "__main__":
    main()
