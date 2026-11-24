
from datetime import datetime

from openpyxl import load_workbook

from tools.create_excel import OUTPUT_DIR


def list_excel_files() -> dict:
    # List all available Excel workbooks

    files = []
    unreadable_files = []

    try:
        # A missing directory is treated as an empty directory.
        if not OUTPUT_DIR.exists():
            return {
                "status": "success",
                "total_files": 0,
                "files": [],
                "unreadable_files": [],
                "message": "No workbook directory exists yet.",
            }

        if not OUTPUT_DIR.is_dir():
            raise NotADirectoryError(
                f"'{OUTPUT_DIR}' is not a directory."
            )

        for file_path in OUTPUT_DIR.iterdir():

            # Ignore directories, temporary files and non-xlsx files.
            if (
                not file_path.is_file()
                or file_path.name.startswith("~$")
                or file_path.suffix.lower() != ".xlsx"
            ):
                continue

            try:
                stat = file_path.stat()

                # Verify that the file is a readable Excel workbook.
                workbook = load_workbook(
                    file_path,
                    read_only=True,
                    data_only=False,
                )
                sheet_names = workbook.sheetnames
                workbook.close()

                files.append({
                    "filename": file_path.name,
                    "size_bytes": stat.st_size,
                    "last_modified": datetime.fromtimestamp(
                        stat.st_mtime
                    ).isoformat(timespec="seconds"),
                    "worksheets": sheet_names,
                })

            except Exception as error:
                unreadable_files.append({
                    "filename": file_path.name,
                    "reason": str(error),
                })

        files.sort(
            key=lambda item: item["filename"].lower()
        )

        return {
            "status": "success",
            "total_files": len(files),
            "files": files,
            "unreadable_files": unreadable_files,
            "message": (
                "No readable Excel workbooks found."
                if not files
                else "Excel workbooks retrieved successfully."
            ),
        }

    except OSError as error:
        return {
            "status": "error",
            "message": f"Unable to access workbook directory: {error}",
        }
