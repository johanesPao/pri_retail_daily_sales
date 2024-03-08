# syntax=docker/dockerfile:1
# escape=

# BUILD IMAGE menggunakan python image dengan tag alpine
FROM python:alpine AS build
WORKDIR /rdsr
COPY . .
RUN \
    # binutils dengan objdump dibutuhkan oleh pyinstaller  dan upx untuk kompressing executable
    apk add binutils upx && \
    # install virtualenv
    pip install virtualenv && \
    # buat virtualenv
    python -m venv rds && \
    # aktivasi virtual environment rds
    source rds/bin/activate && \
    # install modul yang digunakan
    pip install -r requirements.txt && \
    # install pyinstaller untuk compile python
    pip install pyinstaller && \
    # compile python script
    pyinstaller daily_retail_report.py --clean --nowindow --onefile --name pri_rdsr && \
    # kompresi dengan upx
    upx --best /rdsr/dist/pri_rdsr

# FINAL IMAGE menggunakan alpine dengan tag latest (~3 MB)
FROM alpine:latest AS final
WORKDIR /rdsr_service
COPY --from=build /rdsr/dist/pri_rdsr /rdsr_service/pri_rdsr
RUN \
    # install package tzdata
    apk add tzdata && \
    # set timezone pada container dengan softlink /usr/share/zoneinfo ke /etc/localtime
    ln -s /usr/share/zoneinfo/Asia/Jakarta /etc/localtime && \
    # buat direktori untuk output report
    mkdir -p /rdsr_service/reports
CMD ["./pri_rdsr"]