# For Sandbox - code execution and file editing only, no web search or skills

FROM python:3.12-slim

RUN apt-get update && apt-get install -y --no-install-recommends \
    bash curl git jq pandoc poppler-utils \
    libreoffice-writer \
    && rm -rf /var/lib/apt/lists/*

RUN curl -fsSL https://deb.nodesource.com/setup_20.x | bash - \
    && apt-get install -y --no-install-recommends nodejs \
    && rm -rf /var/lib/apt/lists/* \
    && npm install -g docx

RUN pip install --no-cache-dir lxml pyyaml python-docx defusedxml

WORKDIR /workspace
CMD ["sleep", "infinity"]
