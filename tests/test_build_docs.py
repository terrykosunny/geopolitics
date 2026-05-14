"""각 docx 빌드 스크립트가 정상 실행되어 파일을 생성하는지 확인."""

from pathlib import Path
import subprocess
import sys

REPO = Path(__file__).resolve().parent.parent
SCRIPTS = REPO / "scripts"
OUTPUT = REPO / "output"


def _run(script):
    result = subprocess.run(
        [sys.executable, str(SCRIPTS / script)],
        cwd=REPO, capture_output=True, text=True, timeout=60,
    )
    assert result.returncode == 0, f"{script} failed: {result.stderr}"
    path = Path(result.stdout.strip().splitlines()[-1])
    assert path.exists() and path.stat().st_size > 5_000, f"{path} not generated"


def test_demand_letter():
    _run("build_demand_letter_docx.py")


def test_transcript_summary():
    _run("build_transcript_docx.py")


def test_full_transcript():
    _run("build_full_transcript_docx.py")


if __name__ == "__main__":
    test_demand_letter()
    test_transcript_summary()
    test_full_transcript()
    print("ALL OK")
