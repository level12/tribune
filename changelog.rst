Changelog
=========

0.3.0 released 2024-04-23
-------------------------

- configure release via pyp (2da1808_)
- **BC Break** openpyxl support, deprecates use of xlrd (1506f72_)

  - xlrd is unmaintained, and as of version 2, no longer supports reading XLSX format
  - openpyxl is now used for sheet importer functionality (xlsxwriter still used for writing)
  - XlrdSheetImporter may be used instead of SheetImporter for legacy projects
  - xlrd-related classes are deprecated and will be removed in a future release

- **BC Break** update python CI targets, drop python 2 support (ae63e61_)
- update CI config/target, modernize build, resolve deprecations (dc8d5fc_)

.. _2da1808: https://github.com/level12/tribune/commit/2da1808
.. _1506f72: https://github.com/level12/tribune/commit/1506f72
.. _ae63e61: https://github.com/level12/tribune/commit/ae63e61
.. _dc8d5fc: https://github.com/level12/tribune/commit/dc8d5fc


0.2.2 - 2018-02-07
------------------

- Fix deep copy behavior for sheet units using complex tuple expressions and SQLAlchemy objects

0.2.1 - 2017-03-13
------------------

- Fix the 0.2.0 release where the sdist included the whole wheelhouse

0.2.0 - 2017-03-13
------------------

- Fix bug which would cause non-string data in headers to fail (7770e84_)
- Upgrade build / test environment

.. _7770e84: https://github.com/level12/tribune/commit/7770e844aa5e4ded4f926349e6da038c30121809

0.1.1 - 2016
------------

 - Fixed column instance setup when referencing SQLAlchemy attributes and wrapped functions
 - Made SheetImporter easier to construct dynamically at runtime
 - Added parsers for lists, mappings, and nullable fields
