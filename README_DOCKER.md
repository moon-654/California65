# Docker / Dokploy Deployment

This app is a Streamlit service. The container listens on `0.0.0.0:8501`, so it can be reached through Dokploy's proxy or through an internal network port mapping.

## Build locally

```powershell
docker build -t prop65-imds-tool .
```

## Run locally or on an internal server

```powershell
docker run --rm -p 8501:8501 prop65-imds-tool
```

Then open:

```text
http://<server-ip-or-internal-hostname>:8501
```

## Dokploy settings

- Build type: Dockerfile
- Dockerfile path: `Dockerfile`
- Container port: `8501`
- Health path: `/_stcore/health`
- Optional environment variable: `PORT=8501`

After deployment, route the Dokploy domain or internal hostname to container port `8501`.
