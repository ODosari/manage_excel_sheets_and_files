from typing import Literal

from pydantic import Field, field_validator
from pydantic_settings import BaseSettings, SettingsConfigDict


LOG_LEVEL_CHOICES = {"DEBUG", "INFO", "WARN", "ERROR"}
_LOG_LEVEL_ALIASES = {"WARNING": "WARN"}


def normalize_log_level(value: object) -> str:
    """Normalize log-level inputs to the canonical choices used by the CLI."""

    if value is None:
        return "INFO"

    text = str(value).strip()
    if not text:
        return "INFO"

    upper = text.upper()
    return _LOG_LEVEL_ALIASES.get(upper, upper)


class ExcelMgrSettings(BaseSettings):
    model_config = SettingsConfigDict(env_prefix="EXCELMGR_", env_file=".env", extra="ignore")

    # Defaults for CLI options
    glob: str = Field(default="*.xlsx,*.xlsm", description="Default glob patterns")
    recursive: bool = Field(default=False)
    log_format: Literal["json","text"] = Field(default="json")
    log_level: str = Field(default="INFO")
    macro_policy: Literal["warn","forbid","ignore"] = Field(default="warn")
    temp_dir: str | None = Field(default=None, description="Custom temp dir")

    @field_validator("log_level", mode="before")
    @classmethod
    def _normalize_log_level(cls, value: object) -> str:
        return normalize_log_level(value)


settings = ExcelMgrSettings()
