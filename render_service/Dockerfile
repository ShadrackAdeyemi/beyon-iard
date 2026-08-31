FROM python:3.11-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY main.py .
# report_template.docx is never committed to git or copied into this image — it's
# Restricted-classified content (BRD Section 10). It must be provided to the running
# container instead, via a Render Secret File (mounted at /etc/secrets/report_template.docx,
# matching main.py's default TEMPLATE_PATH) or a persistent Disk (set TEMPLATE_PATH env var
# to the mount path). See render_service/README.md "Deploying on Render" for exact steps.
EXPOSE 8000
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
