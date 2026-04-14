"""soffice (LibreOffice) wrapper — convert documents to PDF."""

import os
import subprocess
import sys
import tempfile
from pathlib import Path


SOFFICE_EXTENSIONS = {".doc", ".docx", ".ppt", ".pptx", ".odt", ".odp"}


def _soffice_executable() -> str:
    libre_office_path = os.environ.get("LIBRE_OFFICE_PATH")
    if libre_office_path:
        exe = Path(libre_office_path) / (
            "soffice.exe" if sys.platform == "win32" else "soffice"
        )
        return str(exe)
    return "soffice"


class SofficeWrapper:
    """Convert documents to PDF using soffice (LibreOffice)."""

    def convert(self, input_file: Path, output_file: Path) -> None:
        """
        Convert a document to PDF.

        Args:
            input_file: Path to the input document.
            output_dir: Directory for the output PDF.

        Returns:
            Path to the generated PDF.
        """
        output_file.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory() as tmp_profile:
            profile_url = "file:///" + tmp_profile.replace("\\", "/")
            result = subprocess.run(
                [
                    _soffice_executable(),
                    f"-env:UserInstallation={profile_url}",
                    "--headless",
                    "--invisible",
                    "--norestore",
                    "--nofirststartwizard",
                    "--nologo",
                    "--nolockcheck",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    str(output_file.parent),
                    str(input_file),
                ],
                capture_output=True,
                stdin=subprocess.DEVNULL,
                timeout=120,
                creationflags=(
                    subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
                ),
            )

        if result.returncode != 0:
            raise ValueError(f"soffice 转换失败: {input_file.name}")

        if not output_file.exists():
            raise ValueError(f"PDF 未生成: {input_file.name}")
