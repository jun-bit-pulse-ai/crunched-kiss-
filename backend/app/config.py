from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file=("../.env", ".env"), extra="ignore")

    anthropic_api_key: str = ""
    anthropic_model: str = "claude-sonnet-5"
    cors_origins: list[str] = [
        "https://localhost:3000",
        "https://127.0.0.1:3000",
    ]


settings = Settings()
