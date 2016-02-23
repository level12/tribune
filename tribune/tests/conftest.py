from decimal import Decimal as D


def pytest_configure(config):
    from sqlalchemy import engine_from_config
    from sqlalchemy.orm import sessionmaker, scoped_session
    engine = engine_from_config({'url': 'sqlite:///'}, prefix='')
    session = scoped_session(sessionmaker(bind=engine))

    from tribune.tests import entities
    entities.meta.bind = engine
    entities.meta.create_all()
    entities.session = session
    for x in range(1, 50):
        p = entities.Person()
        p.firstname = 'fn%03d' % x
        p.lastname = 'ln%03d' % x
        p.sortorder = x
        p.numericcol = D('29.26') * x / D('.9')
        session.add(p)

    session.commit()
