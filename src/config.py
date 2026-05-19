import os
from urllib.parse import urlparse, parse_qs
from dotenv import load_dotenv
from openai import AzureOpenAI

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _parse_azure_url(endpoint_url: str):
    parsed = urlparse(endpoint_url)
    azure_endpoint = f"{parsed.scheme}://{parsed.netloc}/"
    path_parts = parsed.path.strip("/").split("/")
    deployment = path_parts[path_parts.index("deployments") + 1]
    api_version = parse_qs(parsed.query).get("api-version", ["2025-01-01-preview"])[0]
    return azure_endpoint, deployment, api_version


def load_config():
    load_dotenv(os.path.join(BASE_DIR, ".env"))

    endpoint_url = os.environ["OPEN_AI_LLM_ENDPOINT"]
    api_key = os.environ["OPEN_AI_LLM_KEY"]
    azure_endpoint, deployment, api_version = _parse_azure_url(endpoint_url)

    client = AzureOpenAI(
        azure_endpoint=azure_endpoint,
        api_key=api_key,
        api_version=api_version,
    )

    # Optional dedicated embedding endpoint (separate Azure deployment).
    # If OPEN_AI_EMBEDDING_ENDPOINT is set, we build a second AzureOpenAI client
    # bound to that deployment. Otherwise embeddings degrade to local fallback.
    embedding_client = None
    embedding_deployment = None
    emb_endpoint_url = os.environ.get("OPEN_AI_EMBEDDING_ENDPOINT")
    if emb_endpoint_url:
        emb_key = os.environ.get("OPEN_AI_EMBEDDING_KEY", api_key)
        emb_endpoint, emb_deployment, emb_api_version = _parse_azure_url(emb_endpoint_url)
        embedding_client = AzureOpenAI(
            azure_endpoint=emb_endpoint,
            api_key=emb_key,
            api_version=emb_api_version,
        )
        embedding_deployment = emb_deployment

    return {
        "client": client,
        "deployment": deployment,
        "embedding_client": embedding_client,
        "embedding_deployment": embedding_deployment,
        "datalake_dir": os.path.join(BASE_DIR, "datalake"),
        "input_dir": os.path.join(BASE_DIR, "sampleinput"),
        "output_dir": os.path.join(BASE_DIR, "output"),
        # MCP mode: default template used when caller does not pick one.
        "default_template": os.environ.get(
            "DEFAULT_TEMPLATE", "Healthcare Client Proposal.pptx"
        ),
    }
