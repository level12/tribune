from blazeutils.strings import randchars
import sqlalchemy as sa
from sqlalchemy.ext.declarative import declarative_base

from .utils import DefaultMixin

meta = sa.MetaData()
session = None
Base = declarative_base(metadata=meta)


class Person(Base, DefaultMixin):
    __tablename__ = 'persons'

    id = sa.Column(sa.Integer, primary_key=True)
    firstname = sa.Column(sa.String(50))
    lastname = sa.Column('last_name', sa.String(50))
    intcol = sa.Column(sa.Integer)
    floatcol = sa.Column(sa.Float)

    def __repr__(self):
        return '<Person: "%s, created: %s">' % (self.id, self.createdts)

    @classmethod
    def testing_create(cls, **kwargs):
        kwargs['firstname'] = kwargs.get('firstname') or randchars()
        return cls.add(**kwargs)
