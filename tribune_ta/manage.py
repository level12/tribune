import urllib

import flask
from flask.ext.script import Manager, Command, Option

from kegreports_ta.app import create_app
import kegreports_ta.model as model
from kegreports_ta.model.helpers import clear_db


class CreateDB(Command):
    """ create the DB and fill with data, optionally clear first """

    option_list = (
        Option('--clear', default=False, dest='clear', action="store_true",
               help='DROP all DB objects first'),
    )

    def run(self, clear):
        if clear:
            clear_db()
            print '- db cleared'

        model.load_db()
        print '- db loaded'


manager = Manager(create_app)
manager.add_option('-c', dest='config', default='Dev',
                   help='flask configuration to use', required=False)
manager.add_command('create-db', CreateDB())


@manager.command
def list_routes(name='list-routes'):
    output = []
    for rule in flask.current_app.url_map.iter_rules():
        methods = ','.join(rule.methods)
        line = urllib.unquote("{:50s} {:20s} {}".format(rule.endpoint, methods, rule))
        output.append(line)

    for line in sorted(output):
        print(line)


@manager.command
def serve():
    flask.current_app.run(debug=True)


def script_entry():
    manager.run()

if __name__ == "__main__":
    script_entry()
