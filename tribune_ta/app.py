import warnings

from flask import Flask
from flask.ext.bootstrap import Bootstrap

# ignore warning about Decimal lossy conversion with SQLite from SA
warnings.filterwarnings('ignore', '.*support Decimal objects natively.*')


def create_app(config):
    app = Flask(__name__)

    app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///'
    #app.config['SQLALCHEMY_ECHO'] = True
    #app.config['DEBUG'] = True
    if config == 'Test':
        app.config['TEST'] = True
    app.secret_key = 'only-testing'

    from tribune_ta.model import db
    db.init_app(app)
    Bootstrap(app)

    return app
