from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from app.routers import process, templates

app = FastAPI(title="Zoho Master Report Processor")

app.add_middleware(
    CORSMiddleware,
    # Local dev (any port) + any Netlify deploy (production and preview
    # subdomains both end in .netlify.app).
    allow_origin_regex=r"http://(localhost|127\.0\.0\.1):\d+|https://[\w-]+\.netlify\.app",
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(templates.router)
app.include_router(process.router)


@app.get("/api/health")
def health():
    return {"status": "ok"}
