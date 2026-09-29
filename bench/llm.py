"""OpenAI client for the Azure AI Foundry project, signed in with Entra ID.

First use prints a device code (sign in at the URL shown). The token is cached in the macOS
Keychain and the account record in ~/.excel-agent/, so later runs are silent.
    uv run python bench/llm.py            # sign in + smoke test

Settings come from environment variables or a `.env` file in the project root (git-ignored; copy
`.env.example`): AZURE_OPENAI_ENDPOINT, AZURE_AI_PROJECT_ENDPOINT, AZURE_TENANT_ID.
"""
import os
import time
from pathlib import Path

from azure.identity import (AuthenticationRecord, AuthenticationRequiredError, DeviceCodeCredential,
                            TokenCachePersistenceOptions, get_bearer_token_provider)
from openai import OpenAI, Timeout

def _load_env(path: Path = Path(__file__).resolve().parent.parent / ".env") -> None:
    """Minimal .env reader (KEY=value lines); real environment variables take precedence."""
    if path.exists():
        for line in path.read_text(encoding="utf-8").splitlines():
            key, sep, value = line.strip().partition("=")
            if sep and key and not key.startswith("#"):
                os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


_load_env()
# e.g. https://<resource>.services.ai.azure.com/openai/v1
ENDPOINT = os.environ.get("AZURE_OPENAI_ENDPOINT", "")
# e.g. https://<resource>.services.ai.azure.com/api/projects/<project>
PROJECT_ENDPOINT = os.environ.get("AZURE_AI_PROJECT_ENDPOINT", "")
SCOPE = "https://ai.azure.com/.default"
# Personal Microsoft accounts must sign in to the subscription's own tenant (e.g. <name>.onmicrosoft.com),
# not /common. Leave empty for work accounts.
TENANT_ID = os.environ.get("AZURE_TENANT_ID", "")
DEFAULT_MODEL = "gpt-5-nano"
# A call that never answers must not hold its job: the SDK's default waits 10 minutes a try, three tries. The
# longest real call seen took about 6 minutes (a reviewer running on to its output limit). The SDK retries a timeout.
READ_TIMEOUT = float(os.environ.get("LLM_READ_TIMEOUT") or 300)
MAX_RETRIES = int(os.environ.get("LLM_MAX_RETRIES") or 2)
RECORD = Path.home() / ".excel-agent" / "auth_record.json"
CACHE = TokenCachePersistenceOptions(name="excel-agent")


def _show_code(verification_uri, user_code, expires_on):
    print(f"DEVICE CODE: go to {verification_uri} and enter {user_code}", flush=True)


def credential(interactive: bool = True) -> DeviceCodeCredential:
    """interactive=False raises AuthenticationRequiredError instead of printing a new device code."""
    record = AuthenticationRecord.deserialize(RECORD.read_text(encoding="utf-8")) if RECORD.exists() else None
    cred = DeviceCodeCredential(tenant_id=TENANT_ID or None, prompt_callback=_show_code, timeout=900,
                                cache_persistence_options=CACHE, authentication_record=record,
                                disable_automatic_authentication=not interactive)
    if record is None:
        if not interactive:
            raise AuthenticationRequiredError(scopes=[SCOPE], message="not signed in")
        RECORD.parent.mkdir(exist_ok=True)
        RECORD.write_text(cred.authenticate(scopes=[SCOPE]).serialize(), encoding="utf-8")
    return cred


def client(interactive: bool = True) -> OpenAI:
    if not ENDPOINT:
        raise RuntimeError("Set AZURE_OPENAI_ENDPOINT (see .env.example)")
    return OpenAI(base_url=ENDPOINT, api_key=get_bearer_token_provider(credential(interactive), SCOPE),
                  timeout=Timeout(READ_TIMEOUT, connect=10.0), max_retries=MAX_RETRIES)


def create(llm: OpenAI, model: str, on_wait=None, purpose: str | None = None, log: dict | None = None, **kwargs):
    """responses.create, kept a safety margin below the deployment's rate limits (see ratelimit.py), and logged
    with what it was for (calllog.py: purpose, plus log= tags or the running job's). on_wait(seconds, limits) is
    called if the call has to wait for room; that wait isn't counted in the call's time."""
    import calllog
    import ratelimit
    est = ratelimit.estimate_tokens(kwargs.get("instructions", ""), kwargs.get("input", ""), kwargs.get("tools", ""),
                                    reserve_output=kwargs.get("max_output_tokens") or 1000)
    entry = ratelimit.acquire(model, est, on_wait)
    r, err, t0 = None, None, time.time()
    try:
        r = llm.responses.create(model=model, **kwargs)
        return r
    except Exception as e:
        err = e
        raise
    finally:
        u = getattr(r, "usage", None)
        ratelimit.settle(entry, (u.input_tokens + u.output_tokens) if u else est)
        calllog.record(model, kwargs, r, time.time() - t0, err, purpose, log)


if __name__ == "__main__":
    r = client().responses.create(model=DEFAULT_MODEL, input="What is the capital of France?")
    print("answer:", r.output_text)
