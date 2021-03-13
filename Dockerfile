FROM python:3.8
LABEL maintainer "Wim van der Ham <wfjvdham@gmail.com>"

ENV PYTHONUNBUFFERED True

RUN mkdir /code
WORKDIR /code

COPY requirements.txt /code/
RUN pip install -r requirements.txt
COPY . /code/

EXPOSE 5050

CMD ["python", "./app.py"]