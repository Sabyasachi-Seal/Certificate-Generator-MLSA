# Use the official Python image as the base image
FROM ubuntu:22.04

ENV DEBIAN_FRONTEND noninteractive

RUN apt-get update && apt-get -y upgrade && \
    apt-get -y install python3.9 && \
    apt update && apt install python3-pip -y

# Method1 - installing LibreOffice and java
RUN apt-get install -y build-essential libssl-dev libffi-dev python3-dev
RUN apt-get --no-install-recommends install libreoffice -y
RUN apt-get install -y libreoffice-java-common

WORKDIR /app

# Copy the requirements file into the container at /app
COPY requirements.txt /app/

# Install any needed packages specified in requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# Copy the current directory contents into the container at /app
COPY . /app/

# Expose the port that the app will run on
EXPOSE 8000

ENV LC_ALL=C.UTF-8
ENV LANG=C.UTF-8

# Command to run your application
CMD ["python3", "main_certificate.py"]
