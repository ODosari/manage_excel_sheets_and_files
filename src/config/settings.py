from pydantic_settings import BaseSettings, SettingsConfigDict
from pydantic import Field
from typing import Literal, Optional

class ExcelMgrSettings(BaseSettings):
    model_config = SettingsConfigDict(env_prefix="EXCELMGR_", env_file=".env", extra="ignore")

    # Defaults for CLI options
    glob: str = Field(default="*.xlsx,*.xlsm", description="Default glob patterns")
    recursive: bool = Field(default=False)
    log_format: Literal["json","text"] = Field(default="json")
    log_level: str = Field(default="INFO")
    macro_policy: Literal["warn","forbid","ignore"] = Field(default="warn")
    temp_dir: Optional[str] = Field(default=None, description="Custom temp dir")

settings = ExcelMgrSettings()
