# app/__init__.py

from flask import Flask

def create_app():
    app = Flask(__name__)

    # Initialize any extensions or configurations
    # e.g., app.config.from_object('config.Config')

    return app
