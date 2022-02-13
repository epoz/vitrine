FROM ubuntu:focal
ARG DEBIAN_FRONTEND=noninteractive
RUN apt-get update -y && apt-get upgrade -y && apt-get -y install python3 python3-pip curl
RUN curl -o- https://raw.githubusercontent.com/nvm-sh/nvm/v0.39.1/install.sh | bash
WORKDIR /out
COPY package.json .
COPY blog.py .
COPY templates .
COPY src .
COPY .envrc .
RUN nvm install --lts
RUN npm install
RUN npm run build
RUN pip install -r requirements.txt
RUN source .envrc 
RUN python blog.py