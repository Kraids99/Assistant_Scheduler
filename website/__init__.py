from flask import Flask
import secrets

def create_app():
    app = Flask(__name__)
    from .urls import urls
    app.register_blueprint(urls)

    return app
