from flask import Flask

def create_app():
    app = Flask(__name__)
    from .urls import urls
    app.register_blueprint(urls)

    return app
