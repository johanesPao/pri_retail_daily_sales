# syntax=docker/dockerfile:1
# escape=

FROM python:slim
WORKDIR /rdsr
COPY . .
RUN pip install -r requirements.txt

