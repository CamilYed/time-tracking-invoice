FROM ubuntu:20.04

RUN echo "Etc/UTC" > /etc/timezone && \
    apt-get update && \
    DEBIAN_FRONTEND=noninteractive apt-get install -y tzdata


ENV TZ=UTC
RUN apt-get install -y tzdata


RUN sed -i 's|http://archive.ubuntu.com|http://ftp.ubuntu.com|g' /etc/apt/sources.list && \
    apt-get update --fix-missing && \
    apt-get install -y \
    wine64 \         
    python3 \
    python3-pip \
    python3-dev \
    build-essential \
    curl \
    && rm -rf /var/lib/apt/lists/*


RUN pip3 install pyinstaller


ENV WINEARCH=win64
ENV WINEPREFIX=/wine
