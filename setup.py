import os.path as osp
try:
    from setuptools import setup
except ImportError:
    from ez_setup import use_setuptools
    use_setuptools()
    from setuptools import setup

# pip install -e .[develop]
develop_requires = [
    'SQLAlchemy',
    'mock',
    'pytest',
    'Flask',
    'Flask-Bootstrap',
    'Flask-SQLAlchemy',
    'Flask-WebTest',
    'wrapt',
    'xlrd',
]

cdir = osp.abspath(osp.dirname(__file__))
README = open(osp.join(cdir, 'readme.rst')).read()
CHANGELOG = open(osp.join(cdir, 'changelog.rst')).read()
VERSION = open(osp.join(cdir, 'tribune', 'version.txt')).read().strip()

setup(
    name="tribune",
    version=VERSION,
    description="A library for coding Excel reports in a declarative fashion",
    long_description='\n\n'.join((README, CHANGELOG)),
    author="Matt Lewellyn",
    author_email="matt.lewellyn@level12.io",
    url='https://bitbucket.org/level12/tribune',
    classifiers=[
        'Development Status :: 4 - Beta',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: BSD License',
        'Operating System :: OS Independent',
        'Programming Language :: Python :: 2.7',
    ],
    license='BSD',
    packages=['tribune'],
    extras_require={'develop': develop_requires},
    zip_safe=False,
    include_package_data=True,
    install_requires=[
        'BlazeUtils',
        'six',
        'xlsxwriter',
    ],
    entry_points="""
        [console_scripts]
        tribune_ta = tribune_ta.manage:script_entry
    """,
)
