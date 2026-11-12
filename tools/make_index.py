"""A PEP 503 simple index of every release's wheels and sdists: site/simple/index.html and
site/simple/sheet-reader/index.html, linking the GitHub release assets with their sha256.

    python3 tools/make_index.py <owner/repo> <output folder>      (needs the gh CLI, authenticated)
"""

import hashlib
import html
import json
import subprocess
import sys
import urllib.request
from pathlib import Path


def assets(repo: str) -> list[tuple[str, str]]:
    out = subprocess.run(["gh", "api", "--paginate", f"repos/{repo}/releases"], capture_output=True, text=True, check=True)
    found = []
    for release in json.loads(out.stdout.replace("][", ",")):
        for asset in release["assets"]:
            if asset["name"].endswith((".whl", ".tar.gz")):
                found.append((asset["name"], asset["browser_download_url"]))
    return sorted(found)


def sha256(url: str) -> str:
    with urllib.request.urlopen(url) as response:  # noqa: S310 (GitHub release asset)
        return hashlib.sha256(response.read()).hexdigest()


def main(repo: str, out: Path) -> None:
    package = out / "simple" / "sheet-reader"
    package.mkdir(parents=True, exist_ok=True)
    (out / "simple" / "index.html").write_text(
        '<!DOCTYPE html><html><body><a href="sheet-reader/">sheet-reader</a></body></html>\n', encoding="utf-8")
    links = "".join(f'<a href="{html.escape(url)}#sha256={sha256(url)}">{html.escape(name)}</a><br>\n'
                    for name, url in assets(repo))
    (package / "index.html").write_text(f"<!DOCTYPE html><html><body>\n{links}</body></html>\n", encoding="utf-8")
    print(f"{links.count('<a ')} files")


if __name__ == "__main__":
    main(sys.argv[1], Path(sys.argv[2]))
