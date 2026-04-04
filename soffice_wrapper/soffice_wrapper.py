"""soffice (LibreOffice) wrapper — convert documents to PDF."""

import subprocess
import sys
import tempfile
from pathlib import Path


SOFFICE_EXTENSIONS = {".doc", ".docx", ".ppt", ".pptx", ".odt", ".odp"}


class SofficeWrapper:
    """Convert documents to PDF using soffice (LibreOffice)."""

    def convert(self, input_file: Path, output_dir: Path) -> Path:
        """
        Convert a document to PDF.

        Args:
            input_file: Path to the input document.
            output_dir: Directory for the output PDF.

        Returns:
            Path to the generated PDF.
        """
        output_dir.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory() as tmp_profile:
            profile_url = "file:///" + tmp_profile.replace("\\", "/")
            result = subprocess.run(
                [
                    "soffice",
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
                    str(output_dir),
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

        output_pdf = output_dir / (input_file.stem + ".pdf")
        if not output_pdf.exists():
            raise ValueError(f"PDF 未生成: {input_file.name}")

        return output_pdf
