import sys, os, time
sys.path.insert(0, r"C:\Users\MP946KB\WORKDIR\Proposal-assist\ProposalAssistant\src")
sys.path.insert(0, r"C:\Users\MP946KB\WORKDIR\Proposal-assist\ProposalAssistant")
os.environ.setdefault("AZURE_OPENAI_API_KEY", "dummy")
os.environ.setdefault("AZURE_OPENAI_ENDPOINT", "https://dummy/")
os.environ.setdefault("AZURE_OPENAI_DEPLOYMENT", "dummy")
from app import app
c = app.test_client()
for tpl_url, tpl in [
    ("Oil%20%26%20Gas%20Sector.pptx", "Oil & Gas Sector.pptx"),
    ("Technology%20Strategy.pptx",    "Technology Strategy.pptx"),
]:
    t0 = time.time()
    r = c.get(f"/templates/layouts/{tpl_url}")
    print(f"\n=== {tpl} in {time.time()-t0:.1f}s ===")
    real = fb = 0
    for lay in (r.get_json() or {}).get("layouts", []):
        path = os.path.join(r"C:\Users\MP946KB\WORKDIR\Proposal-assist\ProposalAssistant", lay["preview_url"].lstrip("/"))
        sz = os.path.getsize(path) if os.path.isfile(path) else 0
        if sz > 30000: real += 1
        elif sz > 0: fb += 1
        print(f"  {sz:>7d} bytes  [{lay['index']:02d}] {lay['name']}")
    print(f"  -- real={real} fb={fb}")
