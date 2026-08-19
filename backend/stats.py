"""Thống kê truy cập webapp (access log JSONL).

Ghi mỗi request quan trọng 1 dòng JSON vào file access.log:
    {"ts": iso, "ip": "...", "method": "GET", "path": "/", "status": 200, "ua": "..."}

Giữ dữ liệu tối đa STATS_RETENTION_DAYS (mặc định 90 ngày) — tự dọn khi ghi.
Không chứa dữ liệu file upload, không chứa nội dung CSV.
"""

import json
import os
import re
from datetime import datetime, timedelta
from pathlib import Path


class Stats:
    """Quản lý access log dạng JSONL + tổng hợp nhanh."""

    def __init__(self, log_dir: Path, retention_days: int = 90, filename: str = "access.log"):
        self.log_dir = Path(log_dir)
        self.log_dir.mkdir(parents=True, exist_ok=True)
        self.log_path = self.log_dir / filename
        self.retention_days = retention_days

    # ------------------------------------------------------------------
    def record(
        self,
        ip: str,
        method: str,
        path: str,
        status: int,
        ua: str = "",
        extra: dict | None = None,
    ) -> None:
        """Ghi 1 dòng log. Dọn log cũ sau khi ghi (thỉnh thoảng)."""
        entry = {
            "ts": datetime.now().isoformat(timespec="seconds"),
            "ip": ip or "",
            "method": method,
            "path": path.split("?")[0],  # không lưu query string
            "status": status,
            "ua": (ua or "")[:200],
        }
        if extra:
            entry.update(extra)
        try:
            with self.log_path.open("a", encoding="utf-8") as f:
                f.write(json.dumps(entry, ensure_ascii=False) + "\n")
        except OSError:
            pass  # không làm chết request nếu không ghi được log

        # Dọn log cũ (mỗi lần ghi — đủ nhẹ vì chỉ quét 1 file).
        self.cleanup()

    # ------------------------------------------------------------------
    def _iter_records(self):
        """Đọc từng dòng JSON hợp lệ trong file log."""
        if not self.log_path.exists():
            return
        with self.log_path.open("r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    yield json.loads(line)
                except json.JSONDecodeError:
                    continue

    # ------------------------------------------------------------------
    def read(self) -> list[dict]:
        return list(self._iter_records())

    # ------------------------------------------------------------------
    def cleanup(self) -> None:
        """Xóa dòng log cũ hơn retention_days (ghi lại file nếu cần)."""
        if not self.log_path.exists():
            return
        cutoff = datetime.now() - timedelta(days=self.retention_days)
        lines = self.log_path.read_text(encoding="utf-8").splitlines()
        keep = []
        for line in lines:
            line = line.strip()
            if not line:
                continue
            try:
                entry = json.loads(line)
                ts = datetime.fromisoformat(entry.get("ts", ""))
            except (json.JSONDecodeError, ValueError):
                continue  # dòng hỏng -> bỏ
            if ts >= cutoff:
                keep.append(line)
        if len(keep) != len(lines):
            self.log_path.write_text("\n".join(keep) + "\n", encoding="utf-8")

    # ------------------------------------------------------------------
    def summarize(self, days: int = 90) -> dict:
        """Tổng hợp: tổng requests, hôm nay, top IP, số upload, keywords, recent."""
        records = self.read()
        now = datetime.now()
        today = now.date().isoformat()
        cutoff = now - timedelta(days=days)

        total = 0
        today_count = 0
        uploads = 0
        ip_counts: dict[str, int] = {}
        hour_counts: dict[int, int] = {}
        day_counts: dict[str, int] = {}
        keyword_counts: dict[str, int] = {}
        recent: list[dict] = []
        unique_ips: set[str] = set()
        error_count = 0
        durations: list[float] = []

        for r in records:
            try:
                ts = datetime.fromisoformat(r.get("ts", ""))
            except ValueError:
                continue
            if ts < cutoff:
                continue
            total += 1
            date = ts.date().isoformat()
            day_counts[date] = day_counts.get(date, 0) + 1
            if date == today:
                today_count += 1
            ip = r.get("ip", "")
            if ip:
                unique_ips.add(ip)
                ip_counts[ip] = ip_counts.get(ip, 0) + 1
            hour = ts.hour
            hour_counts[hour] = hour_counts.get(hour, 0) + 1
            if int(r.get("status", 200) or 200) != 200:
                error_count += 1
            if r.get("duration_ms") is not None:
                durations.append(float(r["duration_ms"]))
            if r.get("method") == "POST" and r.get("path", "").startswith("/process"):
                uploads += 1
                # Đếm từ khóa user nhập (nếu có) — tách theo \n hoặc ,
                kws = r.get("keywords", "")
                if kws:
                    for kw in re.split(r"[\n,]+", str(kws)):
                        kw = kw.strip()
                        if kw:
                            keyword_counts[kw] = keyword_counts.get(kw, 0) + 1
            # 20 record gần nhất (mọi loại request, trừ /stats /health static)
            recent.append(
                {
                    "ts": r.get("ts", ""),
                    "ip": ip,
                    "method": r.get("method", ""),
                    "path": r.get("path", ""),
                    "status": r.get("status", ""),
                    "keywords": r.get("keywords", ""),
                }
            )

        top_ips = sorted(ip_counts.items(), key=lambda kv: kv[1], reverse=True)[:10]
        top_keywords = sorted(keyword_counts.items(), key=lambda kv: kv[1], reverse=True)[:20]

        success_rate = round((total - error_count) / total * 100, 1) if total else 100.0
        avg_duration_ms = round(sum(durations) / len(durations), 1) if durations else 0

        return {
            "total": total,
            "today": today_count,
            "uploads": uploads,
            "unique_ips": len(unique_ips),
            "error_count": error_count,
            "success_rate": success_rate,
            "avg_duration_ms": avg_duration_ms,
            "top_ips": [{"ip": ip, "count": c} for ip, c in top_ips],
            "top_keywords": [{"keyword": kw, "count": c} for kw, c in top_keywords],
            "keyword_count": len(keyword_counts),
            "by_hour": {str(h): hour_counts.get(h, 0) for h in range(24)},
            "by_day": dict(sorted(day_counts.items())),
            "recent": recent[-20:],
            "retention_days": self.retention_days,
        }