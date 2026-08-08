import json
import logging
import os
import tempfile
import time
from pathlib import Path
from urllib.parse import quote

import requests

from app import APP_DIR, INSTANCE_DIR, DEFAULT_DIVISOR, import_file, sync_model_config_from_json

logging.basicConfig(
    level=os.environ.get("SYNC_LOG_LEVEL", "INFO"),
    format="%(asctime)s %(levelname)s source-sync: %(message)s",
)
log = logging.getLogger("source-sync")

SOURCE_REPO = os.environ.get("AWP_SOURCE_REPO", "wincklers85/CalcolatoreAWP").strip()
SOURCE_BRANCH = os.environ.get("AWP_SOURCE_BRANCH", "main").strip() or "main"
SOURCE_FOLDER = os.environ.get("AWP_SOURCE_FOLDER", "Dati").strip().strip("/")
SYNC_INTERVAL = max(300, int(os.environ.get("AWP_SYNC_INTERVAL_SECONDS", "1800")))
DIVISOR = float(os.environ.get("AWP_COUNTER_DIVISOR", str(DEFAULT_DIVISOR)))
TOKEN = os.environ.get("GITHUB_TOKEN", "").strip()


def headers():
    h = {"User-Agent": "AWP-Analisi-SourceSync", "Accept": "application/vnd.github.raw+json"}
    if TOKEN:
        h["Authorization"] = f"Bearer {TOKEN}"
    return h


def raw_url(path: str) -> str:
    safe_path = "/".join(quote(part) for part in path.split("/"))
    return f"https://raw.githubusercontent.com/{SOURCE_REPO}/{SOURCE_BRANCH}/{safe_path}"


def source_path(name: str) -> str:
    value = str(name or "").strip().lstrip("/")
    if SOURCE_FOLDER and not value.startswith(SOURCE_FOLDER + "/"):
        return f"{SOURCE_FOLDER}/{value}"
    return value


def fetch_bytes(path: str, timeout=90) -> bytes:
    r = requests.get(raw_url(path), headers=headers(), timeout=timeout)
    r.raise_for_status()
    return r.content


def fetch_json(path: str):
    return json.loads(fetch_bytes(path, timeout=45).decode("utf-8-sig"))


def sync_cicloslot():
    remote = source_path("cicloslot.json")
    content = fetch_bytes(remote, timeout=45)
    # app.py currently prefers the repository copy, so refresh both locations.
    for target in (APP_DIR / "cicloslot.json", INSTANCE_DIR / "cicloslot.json"):
        target.write_bytes(content)
    count = sync_model_config_from_json()
    log.info("cicloslot sincronizzato: %s modelli", count)
    return count


def sync_sinottici():
    manifest_path = source_path("manifest.json")
    manifest = fetch_json(manifest_path)
    names = manifest.get("sinottici", []) if isinstance(manifest, dict) else []
    names = [str(x).strip() for x in names if str(x).strip().lower().endswith((".xlsx", ".zip"))]

    stats = {"listed": len(names), "new": 0, "duplicate": 0, "rows": 0, "errors": []}
    if not names:
        log.warning("manifest senza sinottici: %s", manifest_path)
        return stats

    with tempfile.TemporaryDirectory() as td:
        td_path = Path(td)
        for idx, name in enumerate(names, 1):
            remote = source_path(name)
            local = td_path / Path(name).name
            try:
                local.write_bytes(fetch_bytes(remote))
                result = import_file(local, Path(name).name, DIVISOR, source_type="github-auto", register_file=True)
                if result.get("file_duplicate"):
                    stats["duplicate"] += 1
                else:
                    stats["new"] += 1
                    stats["rows"] += int(result.get("inserted", 0))
                if result.get("errors"):
                    stats["errors"].extend(result["errors"])
            except Exception as exc:
                msg = f"{name}: {exc}"
                stats["errors"].append(msg)
                log.warning(msg)

            if idx % 25 == 0 or idx == len(names):
                log.info(
                    "sinottici %s/%s - nuovi=%s duplicati=%s errori=%s",
                    idx, len(names), stats["new"], stats["duplicate"], len(stats["errors"]),
                )

    return stats


def sync_once():
    started = time.time()
    log.info("sincronizzazione da %s/%s (%s)", SOURCE_REPO, SOURCE_FOLDER, SOURCE_BRANCH)
    models = sync_cicloslot()
    stats = sync_sinottici()
    log.info(
        "sync completata in %.1fs: manifest=%s nuovi=%s duplicati=%s righe_nuove=%s modelli=%s errori=%s",
        time.time() - started,
        stats["listed"], stats["new"], stats["duplicate"], stats["rows"], models, len(stats["errors"]),
    )


def main():
    while True:
        try:
            sync_once()
        except Exception:
            log.exception("errore durante la sincronizzazione automatica")
        time.sleep(SYNC_INTERVAL)


if __name__ == "__main__":
    main()
