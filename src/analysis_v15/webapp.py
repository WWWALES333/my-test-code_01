from __future__ import annotations

import argparse
import json
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Dict, List
from urllib.parse import urlparse

from .reporter import build_workbench_pages, write_jsonl
from .review_state import (
    apply_review_decision,
    build_review_audit_log,
    build_review_batch_summaries,
    build_review_learning_candidates,
    build_review_learning_summary,
    load_review_decisions,
    validate_review_submission,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="v1.5 AI 一线情报工作台本地服务")
    parser.add_argument("--data", required=True, help="data/output/insights/v1.5 目录")
    parser.add_argument("--host", default="127.0.0.1", help="监听地址")
    parser.add_argument("--port", type=int, default=8765, help="监听端口")
    return parser.parse_args()


def serve(data_dir: Path, host: str, port: int) -> None:
    app = _WorkbenchServer(data_dir)
    server = ThreadingHTTPServer((host, port), app.handler_class())
    print(f"http://{host}:{port}/overview")
    server.serve_forever()


class _WorkbenchServer:
    def __init__(self, data_dir: Path) -> None:
        self.data_dir = data_dir

    def handler_class(self):
        outer = self

        class Handler(BaseHTTPRequestHandler):
            def do_GET(self) -> None:
                parsed = urlparse(self.path)
                path = parsed.path or "/"
                if path in {"/", "/overview", "/overview.html"}:
                    self._html("overview.html")
                    return
                if path in {"/trends", "/trends.html"}:
                    self._html("trends.html")
                    return
                if path in {"/sales", "/sales.html"}:
                    self._html("sales.html")
                    return
                if path in {"/insights", "/insights.html"}:
                    self._html("insights.html")
                    return
                if path in {"/review", "/review.html"}:
                    self._html("review.html")
                    return
                if path in {"/evidence", "/evidence.html"}:
                    self._html("evidence.html")
                    return
                self.send_error(HTTPStatus.NOT_FOUND)

            def do_POST(self) -> None:
                parsed = urlparse(self.path)
                if parsed.path == "/api/rebuild":
                    result = outer._rebuild()
                    self.send_response(HTTPStatus.OK)
                    self.send_header("Content-Type", "application/json; charset=utf-8")
                    self.end_headers()
                    self.wfile.write(json.dumps({"ok": True, "result": result}, ensure_ascii=False).encode("utf-8"))
                    return
                if parsed.path != "/api/review-decisions":
                    self.send_error(HTTPStatus.NOT_FOUND)
                    return
                length = int(self.headers.get("Content-Length", "0"))
                payload = json.loads(self.rfile.read(length) or b"{}")
                review_dir = outer.data_dir / "review"
                tasks = outer._load_review_tasks()
                task_map = {str(row.get("task_id", "")): row for row in tasks}
                task = task_map.get(str(payload.get("task_id", "")))
                if not task:
                    self.send_error(HTTPStatus.BAD_REQUEST, "invalid task_id")
                    return
                reviewed_fields = dict(payload.get("reviewed_fields", {}))
                learning_fields = dict(payload.get("learning_fields", {}))
                errors = validate_review_submission(task, reviewed_fields, learning_fields)
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
                    reviewed_at=str(payload.get("reviewed_at", "")).strip() or _now_text(),
                    review_comment=str(payload.get("review_comment", "")).strip(),
                    learning_fields=learning_fields,
                )
                decisions = load_review_decisions(Path("/nonexistent"), review_dir=review_dir)
                audit_log = build_review_audit_log(decisions)
                candidates = build_review_learning_candidates(decisions)
                learning_summary = build_review_learning_summary(decisions, candidates)
                batch_summaries = build_review_batch_summaries(outer._load_review_tasks(review_dir=review_dir))
                write_jsonl(review_dir / "review_audit_log.jsonl", audit_log)
                write_jsonl(review_dir / "review_candidates.jsonl", candidates)
                (review_dir / "review_learning_summary.json").write_text(json.dumps(learning_summary, ensure_ascii=False, indent=2), encoding="utf-8")
                (review_dir / "review_batch_summaries.json").write_text(json.dumps(batch_summaries, ensure_ascii=False, indent=2), encoding="utf-8")
                self.send_response(HTTPStatus.OK)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.end_headers()
                self.wfile.write(
                    json.dumps(
                        {
                            "ok": True,
                            "decision": decision,
                            "review_learning_summary": learning_summary,
                            "review_candidates": candidates[:20],
                            "review_batch_summaries": batch_summaries,
                        },
                        ensure_ascii=False,
                    ).encode("utf-8")
                )

            def log_message(self, format: str, *args) -> None:  # noqa: A003
                return

            def _html(self, page_name: str) -> None:
                pages = outer._build_pages()
                content = pages.get(page_name)
                if content is None:
                    self.send_error(HTTPStatus.NOT_FOUND)
                    return
                self.send_response(HTTPStatus.OK)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.end_headers()
                self.wfile.write(content.encode("utf-8"))

        return Handler

    def _build_pages(self) -> Dict[str, str]:
        payload = self._load_payload()
        return build_workbench_pages(
            snapshot=payload["snapshot"],
            trend_cube=payload["trend_cube"],
            trend_explanations=payload["trend_explanations"],
            salesperson_profiles=payload["salesperson_profiles"],
            region_rollups=payload["region_rollups"],
            insight_tree=payload["insight_tree"],
            review_tasks=payload["review_tasks"],
            evidence_index=payload["evidence_index"],
            review_learning_summary=payload["review_learning_summary"],
            review_candidates=payload["review_candidates"],
            review_batch_summaries=payload["review_batch_summaries"],
            interactive_review=True,
        )

    def _load_payload(self) -> Dict[str, object]:
        normalized = self.data_dir / "normalized"
        review_dir = self.data_dir / "review"
        snapshot = _read_json(normalized / "dashboard_snapshot.json", {})
        trend_cube = _read_json(normalized / "trend_cube.json", [])
        trend_explanations = _read_json(normalized / "trend_explanations.json", [])
        salesperson_profiles = _read_jsonl(normalized / "salesperson_profile.jsonl")
        region_rollups = _read_jsonl(normalized / "region_sales_rollup.jsonl")
        insight_tree = _read_json(normalized / "insight_tree.json", {"business_lines": []})
        evidence_index = _read_jsonl(normalized / "evidence_index.jsonl")
        review_learning_summary = _read_json(review_dir / "review_learning_summary.json", {})
        review_candidates = _read_jsonl(review_dir / "review_candidates.jsonl")
        review_batch_summaries = _read_json(review_dir / "review_batch_summaries.json", [])
        review_tasks = self._load_review_tasks(review_dir=review_dir)
        return {
            "snapshot": snapshot,
            "trend_cube": trend_cube,
            "trend_explanations": trend_explanations,
            "salesperson_profiles": salesperson_profiles,
            "region_rollups": region_rollups,
            "insight_tree": insight_tree,
            "review_tasks": review_tasks,
            "evidence_index": evidence_index,
            "review_learning_summary": review_learning_summary,
            "review_candidates": review_candidates,
            "review_batch_summaries": review_batch_summaries,
        }

    def _rebuild(self) -> Dict[str, object]:
        manifest = _read_json(self.data_dir / "run_manifest.json", {})
        if not manifest:
            raise RuntimeError("缺少 run_manifest.json，无法自动重建")
        from .run import run_pipeline

        result = run_pipeline(
            samples_dir=Path(str(manifest.get("samples", ""))),
            annotations_dir=Path(str(manifest.get("annotations", ""))),
            out_dir=Path(str(manifest.get("out", str(self.data_dir)))),
            model_mode=str(manifest.get("model_mode", "mock")),
            llm_concurrency=int(manifest.get("llm_concurrency", 4) or 4),
            llm_chunk_size=int(manifest.get("llm_chunk_size", 50) or 50),
            roster_path=Path(str(manifest.get("roster", ""))) if str(manifest.get("roster", "")).strip() else None,
        )
        payload = self._load_payload()
        result["total_ai_mentions"] = int(payload["snapshot"].get("total_ai_mentions", 0))
        result["open_review_tasks"] = int(payload["snapshot"].get("open_review_tasks", 0))
        return result

    def _load_review_tasks(self, review_dir: Path | None = None) -> List[Dict[str, object]]:
        normalized = self.data_dir / "normalized"
        tasks = _read_jsonl(normalized / "review_tasks.jsonl")
        decisions = load_review_decisions(Path("/nonexistent"), review_dir=review_dir or (self.data_dir / "review"))
        task_map = {str(row.get("task_id", "")): dict(row) for row in tasks}
        for row in task_map.values():
            key = (str(row.get("report_id", "")), str(row.get("segment_id", "")))
            decision = decisions.get(key)
            if not decision:
                continue
            row["task_status"] = "reviewed"
            row["edited_fields"] = dict(decision.get("final_labels", {}))
            row["review_comment"] = str(decision.get("review_comment", ""))
            row["reviewer"] = str(decision.get("reviewer", ""))
            row["reviewed_at"] = str(decision.get("reviewed_at", ""))
        return sorted(task_map.values(), key=lambda item: (str(item.get("task_status", "")), str(item.get("task_id", ""))))


def _read_json(path: Path, default):
    if not path.exists():
        return default
    return json.loads(path.read_text(encoding="utf-8"))


def _read_jsonl(path: Path) -> List[Dict[str, object]]:
    if not path.exists():
        return []
    rows: List[Dict[str, object]] = []
    for line in path.read_text(encoding="utf-8").splitlines():
        if line.strip():
            rows.append(json.loads(line))
    return rows


def _now_text() -> str:
    from datetime import datetime

    return datetime.now().isoformat(timespec="seconds")


def main() -> None:
    args = parse_args()
    serve(Path(args.data), args.host, args.port)


if __name__ == "__main__":
    main()
