FROM python:3.11-bookworm
ENV DEBIAN_FRONTEND=noninteractive

RUN apt update && apt install -y python3 python3-pip python3-dev build-essential xvfb gcc libjpeg-dev zlib1g-dev tk-dev libffi-dev libssl-dev gnome-screenshot xdotool firefox-esr xclip x11vnc xorg fluxbox imagemagick && apt clean && rm -rf /var/lib/apt/lists/*

RUN pip install --upgrade pip setuptools wheel
WORKDIR /app
COPY . .
COPY requirements.txt /tmp/requirements.txt
RUN pip install -r /tmp/requirements.txt
RUN apt update && apt install -y locales && locale-gen pt_BR.UTF-8
ENV LANG=pt_BR.UTF-8
ENV LC_ALL=pt_BR.UTF-8

RUN touch /root/.Xauthority
CMD ["sh", "execute.sh" ]