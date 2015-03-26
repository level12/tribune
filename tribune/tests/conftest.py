def pytest_configure(config):
    from tribune_ta.app import create_app
    app = create_app(config='Test')
    app.test_request_context().push()

    from tribune_ta.model import load_db
    load_db()
