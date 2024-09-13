# Use a Python 3.9 slim base image
FROM python:3.9-slim

# Install system dependencies and Java (for tabula-py)
RUN apt-get update && \
    apt-get install -y openjdk-17-jdk build-essential gcc libpq-dev && \
    apt-get clean && rm -rf /var/lib/apt/lists/*

# Set JAVA_HOME and update PATH for Java usage
ENV JAVA_HOME /usr/lib/jvm/java-17-openjdk-amd64
ENV PATH $PATH:$JAVA_HOME/bin

# Set working directory
WORKDIR /app

# Copy the requirements.txt and upgrade pip
COPY requirements.txt .
RUN pip install --upgrade pip && pip install --no-cache-dir -r requirements.txt

# Download the SpaCy language model 'en_core_web_sm'
RUN python -m spacy download en_core_web_sm

# Copy the rest of the application code to the working directory
COPY . .

# Command to run the application using gunicorn
CMD ["gunicorn", "app.main:app", "--bind", "0.0.0.0:5000", "--timeout", "1200"]
