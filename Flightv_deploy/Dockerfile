FROM python:3.12-slim

WORKDIR /app

COPY Flightv_deploy/requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY Flightv_deploy/ .

EXPOSE 10000

CMD ["bash", "start.sh"]
