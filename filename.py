from __future__ import annotations

import argparse
import json
import logging
import os
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Optional, Sequence

MAX_FILENAME_LEN_DEFAULT = 64
MIN_FILENAME_LEN = 1
MAX_FILENAME_LEN = 255  # filesystem sanity bound


@dataclass(frozen=True)
class NameViolation:
    """Represents a single file-name length violation."""
    path: Path
    name_length: int
    redacted_path: str


def _get_logger(logger: Optional[logging.Logger] = None) -> logging.Logger:
    if logger is not None:
        return logger
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    return logging.getLogger(__name__)


def _redact_root(path: Path, root: Path) -> str:
    try:
        rel = path.relative_to(root)
        return f"{root.name}/.../{rel}"
    except ValueError:
        return path.name


def _is_within_root(path: Path, root: Path) -> bool:
    try:
        path.relative_to(root)
        return True
    except ValueError:
        return False


def scan_directory(
    root: Path,
    *,
    max_len: int = MAX_FILENAME_LEN_DEFAULT,
    logger: Optional[logging.Logger] = None,
    trace_id: Optional[str] = None,
) -> List[NameViolation]:
    """
    Scan a directory for files whose base name exceeds `max_len`.

    Args:
        root: Directory to scan.
        max_len: Maximum allowed length for the file name (not full path).
        logger: Optional logger for structured records.
        trace_id: Optional correlation identifier.

    Returns:
        List of NameViolation for all offending files.

    Raises:
        NotADirectoryError: if `root` is not a directory.
        ValueError: if `max_len` is outside sane bounds.
    """
    log = _get_logger(logger)

    if not (MIN_FILENAME_LEN <= max_len <= MAX_FILENAME_LEN):
        raise ValueError(f"max_len must be between {MIN_FILENAME_LEN} and "
                         f"{MAX_FILENAME_LEN}, got {max_len}")

    root_resolved = root.expanduser().resolve(strict=True)
    if not root_resolved.is_dir():
        raise NotADirectoryError(f"Path is not a directory: {root_resolved}")

    violations: List[NameViolation] = []

    for current_root, _, files in os.walk(root_resolved):
        current_root_path = Path(current_root).resolve()
        if not _is_within_root(current_root_path, root_resolved):
            log.warning(
                {
                    "event": "skip_out_of_root_dir",
                    "path": str(current_root_path),
                    "root": str(root_resolved),
                    "trace_id": trace_id,
                }
            )
            continue

        for name in files:
            name_length = len(name)
            if name_length <= max_len:
                continue

            file_path = (current_root_path / name).resolve()
            if not _is_within_root(file_path, root_resolved):
                log.warning(
                    {
                        "event": "skip_out_of_root_file",
                        "path": str(file_path),
                        "root": str(root_resolved),
                        "trace_id": trace_id,
                    }
                )
                continue

            violation = NameViolation(
                path=file_path,
                name_length=name_length,
                redacted_path=_redact_root(file_path, root_resolved),
            )
            violations.append(violation)

    log.info(
        {
            "event": "scan_complete",
            "root": str(root_resolved),
            "violations": len(violations),
            "max_len": max_len,
            "trace_id": trace_id,
        }
    )
    return violations


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Check that file names do not exceed a length limit."
    )
    parser.add_argument("directory", help="Directory to scan.")
    parser.add_argument(
        "--max-len",
        type=int,
        default=MAX_FILENAME_LEN_DEFAULT,
        help="Maximum allowed file-name length (default: 64).",
    )
    parser.add_argument(
        "--trace-id",
        type=str,
        default=None,
        help="Optional trace identifier for log correlation.",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Output violations as JSON lines.",
    )
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)

    logger = _get_logger()

    try:
        root = Path(args.directory)
        violations = scan_directory(
            root,
            max_len=args.max_len,
            logger=logger,
            trace_id=args.trace_id,
        )
    except Exception as exc:
        logger.error(
            {
                "event": "scan_error",
                "error": str(exc),
                "trace_id": args.trace_id,
            }
        )
        return 1

    if violations:
        for v in violations:
            record = {
                "event": "filename_too_long",
                "path": str(v.path),
                "length": v.name_length,
                "redacted_path": v.redacted_path,
                "max_len": args.max_len,
                "trace_id": args.trace_id,
            }
            if args.json:
                logger.warning(json.dumps(record))
            else:
                logger.warning(record)
        return 2

    logger.info(
        {
            "event": "no_violations",
            "trace_id": args.trace_id,
        }
    )
    return 0


def legacy_main() -> None:
    """
    Deprecated entry point kept for drop-in compatibility.
    """
    raise SystemExit(main())


if __name__ == "__main__":
    raise SystemExit(main())
