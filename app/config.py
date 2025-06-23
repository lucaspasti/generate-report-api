from pydantic_settings import BaseSettings
import os
import json

from typing import List


class Settings(BaseSettings):
    SUPABASE_URL: str
    SUPABASE_KEY: str
    ALLOWED_ORIGINS = json.loads(
        os.getenv("ALLOWED_ORIGINS", '["http://localhost:3000"]'))

    class Config:
        env_file = ".env"
        env_file_encoding = "utf-8"


settings = Settings()
