Changelog
---------

0.2.2 - 2018-02-07
==================

- Fix deep copy behavior for sheet units using complex tuple expressions and SQLAlchemy objects

0.2.1 - 2017-03-13
==================

- Fix the 0.2.0 release where the sdist included the whole wheelhouse

0.2.0 - 2017-03-13
==================

- Fix bug which would cause non-string data in headers to fail (7770e84_)
- Upgrade build / test environment

.. 7770e84: https://github.com/level12/tribune/commit/7770e844aa5e4ded4f926349e6da038c30121809

0.1.1 - 2016
============

 - Fixed column instance setup when referencing SQLAlchemy attributes and wrapped functions
 - Made SheetImporter easier to construct dynamically at runtime
 - Added parsers for lists, mappings, and nullable fields

