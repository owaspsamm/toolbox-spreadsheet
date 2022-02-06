FROM python:3.9
COPY . /opt/app
WORKDIR /opt/app
ENV PYTHONPATH /opt/app
RUN pip install -r requirements.txt
CMD ["python", "/opt/app/toolkit_updater.py"]
