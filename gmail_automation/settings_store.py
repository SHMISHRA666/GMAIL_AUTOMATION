from __future__ import annotations

import importlib

from sqlmodel import Session

from .db_models import AppSetting, utc_now_text


KEYRING_SERVICE = "gmail_automation_compliance"


def _load_keyring():
    # Dynamic import keeps keyring optional for packaged runtimes.
    module_name = "key" + "ring"
    try:
        return importlib.import_module(module_name)
    except Exception:
        return None


class SettingsService:
    def __init__(self, session: Session) -> None:
        self.session = session

    def set_value(self, key: str, value: str, secret: bool = False) -> None:
        stored_value = value
        if secret:
            try:
                keyring = _load_keyring()
                if keyring is not None:
                    keyring.set_password(KEYRING_SERVICE, key, value)
                    stored_value = "__keyring__"
            except Exception:
                stored_value = value
        setting = self.session.get(AppSetting, key) or AppSetting(key=key)
        setting.value = stored_value
        setting.is_secret = secret
        setting.updated_at = utc_now_text()
        self.session.add(setting)
        self.session.commit()

    def get_value(self, key: str, default: str = "") -> str:
        setting = self.session.get(AppSetting, key)
        if setting is None:
            return default
        if setting.is_secret and setting.value == "__keyring__":
            try:
                keyring = _load_keyring()
                if keyring is None:
                    return default
                return keyring.get_password(KEYRING_SERVICE, key) or default
            except Exception:
                return default
        return setting.value
