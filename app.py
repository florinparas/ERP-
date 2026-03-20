import os
from flask import Flask
from config import (SECRET_KEY, DATA_DIR, UPLOAD_DIR, SCHEMAS,
                    ANGAJATI_FILE, DEPARTAMENTE_FILE, PONTAJ_FILE,
                    CONCEDII_FILE, EVALUARI_FILE, DOCUMENTE_FILE)
from modules.excel_manager import init_file
from routes.main import main_bp
from routes.angajati import angajati_bp
from routes.departamente import departamente_bp
from routes.pontaj import pontaj_bp
from routes.concedii import concedii_bp
from routes.evaluari import evaluari_bp
from routes.documente import documente_bp


def create_app():
    app = Flask(__name__)
    app.secret_key = SECRET_KEY
    app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max upload

    # Create directories
    os.makedirs(DATA_DIR, exist_ok=True)
    os.makedirs(UPLOAD_DIR, exist_ok=True)

    # Initialize Excel files
    init_file(ANGAJATI_FILE, SCHEMAS['angajati'])
    init_file(DEPARTAMENTE_FILE, SCHEMAS['departamente'])
    init_file(PONTAJ_FILE, SCHEMAS['pontaj'])
    init_file(CONCEDII_FILE, SCHEMAS['concedii'])
    init_file(EVALUARI_FILE, SCHEMAS['evaluari'])
    init_file(DOCUMENTE_FILE, SCHEMAS['documente'])

    # Register blueprints
    app.register_blueprint(main_bp)
    app.register_blueprint(angajati_bp, url_prefix='/angajati')
    app.register_blueprint(departamente_bp, url_prefix='/departamente')
    app.register_blueprint(pontaj_bp, url_prefix='/pontaj')
    app.register_blueprint(concedii_bp, url_prefix='/concedii')
    app.register_blueprint(evaluari_bp, url_prefix='/evaluari')
    app.register_blueprint(documente_bp, url_prefix='/documente')

    return app


if __name__ == '__main__':
    app = create_app()
    app.run(debug=True, host='0.0.0.0', port=5000)
