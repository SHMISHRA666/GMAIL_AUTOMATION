from __future__ import annotations

import os
import platform
from contextlib import contextmanager
from pathlib import Path
from typing import Iterator

from sqlmodel import Session, SQLModel, create_engine

from . import db_models  # noqa: F401 - imported so SQLModel metadata sees all tables


APP_DIR_NAME = "GmailAutomation"
DB_FILE_NAME = "compliance.db"


def app_data_dir() -> Path:
    override = os.environ.get("GMAIL_AUTOMATION_DATA_DIR")
    if override:
        return Path(override).expanduser().resolve()
    system = platform.system().lower()
    if system == "windows":
        base = Path(os.environ.get("APPDATA") or Path.home() / "AppData" / "Roaming")
    elif system == "darwin":
        base = Path.home() / "Library" / "Application Support"
    else:
        base = Path(os.environ.get("XDG_DATA_HOME") or Path.home() / ".local" / "share")
    return base / APP_DIR_NAME


def default_database_path() -> Path:
    return app_data_dir() / DB_FILE_NAME


def database_url(path: Path | None = None) -> str:
    db_path = path or default_database_path()
    db_path.parent.mkdir(parents=True, exist_ok=True)
    return f"sqlite:///{db_path.as_posix()}"


def create_db_engine(path: Path | None = None, echo: bool = False):
    return create_engine(database_url(path), echo=echo, connect_args={"check_same_thread": False})


def init_db(path: Path | None = None, echo: bool = False):
    engine = create_db_engine(path, echo)
    SQLModel.metadata.create_all(engine)
    return engine


@contextmanager
def session_scope(path: Path | None = None) -> Iterator[Session]:
    engine = init_db(path)
    with Session(engine) as session:
        try:
            yield session
            session.commit()
        except Exception:
            session.rollback()
            raise
