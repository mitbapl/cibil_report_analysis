# Use an official Python runtime as a parent image
FROM python:3.9-slim

# Set the working directory in the container
WORKDIR /app

# Copy the current directory contents into the container at /app
COPY . /app

# Install any needed packages specified in requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# Make the 'uploads' directory
RUN mkdir -p /app/uploads

# Expose port 5000 to allow external access to the container
EXPOSE 5000

# Define environment variable for Flask
ENV FLASK_APP=main.py

# Run the application with Gunicorn (for production)
CMD ["gunicorn", "--bind", "0.0.0.0:5000", "main:app"]
