# tribune

[![nox](https://github.com/level12/tribune/actions/workflows/nox.yaml/badge.svg)](https://github.com/level12/tribune/actions/workflows/nox.yaml)

Tribune is a reporting library for Flask and other Python web frameworks. It
allows developers to create Excel report classes with a declarative style
similar to SQLAlchemy's ORM.

The library is currently beta quality. Its source and tests provide the most
complete usage examples.


## Dev


### Copier Template

Project structure and tooling mostly derives from
[Coppy](https://github.com/level12/coppy). See its documentation for context and
additional instructions.

This project can be updated from the upstream repo, see
[Updating a Project](https://github.com/level12/coppy?tab=readme-ov-file#template-updates).


### Project Setup

From zero to hero (passing tests that is):

1. Ensure [host dependencies](https://github.com/level12/coppy/wiki/Mise) are installed

2. Start docker service dependencies (if applicable):

   `docker compose up -d`

3. Sync [project](https://docs.astral.sh/uv/concepts/projects/) virtualenv w/ lock file:

   `uv sync`

4. Configure prek:

   `prek install`

5. Run tests:

   `nox`


### Versions

Versions are date based. A `version` task shows and bumps versions:

```shell

  # Show current version
  mise run version show

  # Bump version based on date, tag, and push:
  mise run version bump

  # See other options
  mise run version -- --help
```
