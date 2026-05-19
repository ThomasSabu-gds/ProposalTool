import sys, os
sys.path.insert(0, r"C:\Users\MP946KB\WORKDIR\Proposal-assist\ProposalAssistant\src")
sys.path.insert(0, r"C:\Users\MP946KB\WORKDIR\Proposal-assist\ProposalAssistant")
os.environ.setdefault("AZURE_OPENAI_API_KEY", "dummy")
os.environ.setdefault("AZURE_OPENAI_ENDPOINT", "https://dummy/")
os.environ.setdefault("AZURE_OPENAI_DEPLOYMENT", "dummy")
from app import app
c = app.test_client()
for tpl in ("Healthcare Client Proposal.pptx","Oil & Gas Sector.pptx","Technology Strategy.pptx"):
    r = c.get(f"/templates/layouts/{tpl}")
    data = r.get_json()
    pngs = []
    for lay in data.get("layouts", []):
        path = os.path.join(r"C:\Users\MP946KB\WORKDIR\Proposal-assist\ProposalAssistant", lay["preview_url"].lstrip("/"))
        size = os.path.getsize(path) if os.path.isfile(path) else 0
        pngs.append((lay["name"], size))
    real = sum(1 for n,s in pngs if s > 30000)
    fb = sum(1 for n,s in pngs if 0 < s <= 30000)
    print(f"{tpl}: total={len(pngs)} real={real} fallback={fb}")
    for nm, sz in pngs[:5]:
        print(f"   {sz:>7d} bytes  {nm}")
