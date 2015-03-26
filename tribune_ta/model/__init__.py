from decimal import Decimal as D

from flask.ext.sqlalchemy import SQLAlchemy

db = SQLAlchemy()


def load_db():
    from tribune_ta.model.entities import Person

    db.create_all()

    for x in range(1, 50):
        p = Person()
        p.firstname = 'fn%03d' % x
        p.lastname = 'ln%03d' % x
        p.sortorder = x
        p.numericcol = D('29.26') * x / D('.9')
        db.session.add(p)

    db.session.commit()
