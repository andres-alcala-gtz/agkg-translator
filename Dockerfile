FROM ubuntu:24.04

ENV DEBIAN_FRONTEND=noninteractive

RUN apt-get update && apt-get install -y \
    python3 \
    python3-pip \
    libreoffice

RUN pip3 install --no-cache-dir --break-system-packages \
    argostranslate==1.11.0 \
    click==8.5.0 \
    openpyxl==3.1.5 \
    pandas==3.1.0 \
    python-docx==1.2.0 \
    python-pptx==1.0.2

WORKDIR /app

COPY ./ ./

CMD [ "python3", "main.py" ]
