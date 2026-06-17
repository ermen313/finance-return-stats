FROM alpine

RUN apk add tzdata && cp /usr/share/zoneinfo/Asia/Shanghai /etc/localtime && echo Asia/Shanghai > /etc/timezone

RUN sed -i 's/dl-cdn.alpinelinux.org/mirrors.tencent.com/g' /etc/apk/repositories \
&& apk add --update --no-cache python3 py3-pip gcc python3-dev linux-headers musl-dev \
&& rm -rf /var/cache/apk/*

RUN apk add ca-certificates

COPY . /app
WORKDIR /app

RUN pip config set global.index-url http://mirrors.cloud.tencent.com/pypi/simple \
&& pip config set global.trusted-host mirrors.cloud.tencent.com \
&& pip install --upgrade pip --break-system-packages \
&& pip install --user -r requirements.txt --break-system-packages

CMD ["python3", "app.py"]
