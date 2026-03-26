"""
FinStatement Projector — FY 2025-26
NSVR & Associates LLP
Developer Version
"""

from flask import Flask
from config import Config


def create_app(config=Config):
    app = Flask(__name__)
    app.config.from_object(config)

    # Register blueprints
    from routes.main   import main_bp
    from routes.api    import api_bp

    app.register_blueprint(main_bp)
    app.register_blueprint(api_bp, url_prefix='/api')

    return app


if __name__ == '__main__':
    app = create_app()
    app.run(debug=True, port=5000)
